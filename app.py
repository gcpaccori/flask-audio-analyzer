from flask import Flask, render_template, request, redirect, url_for, send_from_directory, jsonify
from flask_sqlalchemy import SQLAlchemy
from apscheduler.schedulers.background import BackgroundScheduler
from scipy.io import wavfile
import os
import logging
import numpy as np
import datetime
import json
from flask_socketio import SocketIO, emit
from flask_cors import CORS
import pytz
from flask import send_file


# Configurar logger
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
CORS(app)
socketio = SocketIO(app, cors_allowed_origins="*")
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['API_KEY'] = 'acoustics'
db = SQLAlchemy(app)

# Crear carpeta de subidas si no existe
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])
    logger.debug(f"Carpeta de subidas creada en: {app.config['UPLOAD_FOLDER']}")

# Parámetros acústicos (constantes)
p0 = 20e-6   # 20 µPa, umbral de audición humana
E0 = 4e-10
J0 = 1e-12
I0 = 1e-12
rho = 1.2    # Densidad del aire (kg/m³)
c = 343.0    # Velocidad del sonido (m/s)

# Global para seguimiento del estado de procesamiento de audios
processing_audios = {}

# Umbral de ruido seguro (por ejemplo, 85 dB)
SAFE_NOISE_LEVEL = 85.0

# Función para obtener la hora actual en la zona "America/Lima"
def lima_now():
    lima_tz = pytz.timezone("America/Lima")
    return datetime.datetime.now(lima_tz)

# Modelos de la base de datos

class Audio(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(100), nullable=False)
    title = db.Column(db.String(100), nullable=False)
    source = db.Column(db.String(100), nullable=False)
    # Registra la hora de llegada del audio usando la hora de Lima
    timestamp = db.Column(db.DateTime, default=lima_now)

class Escenario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    descripcion = db.Column(db.String(500), nullable=True)
    ubicacion = db.Column(db.String(100), nullable=True)
    start_time = db.Column(db.DateTime, nullable=False)
    end_time = db.Column(db.DateTime, nullable=False)
    estado = db.Column(db.String(50), nullable=False, default='programado')
    horas_medicion = db.Column(db.Float, nullable=True)
    dias_medicion = db.Column(db.Float, nullable=True)
    tipo_ruido = db.Column(db.String(50), nullable=True)
    num_fuentes = db.Column(db.Integer, nullable=True)
    num_personas = db.Column(db.Integer, nullable=True)
    proteccion_auditiva = db.Column(db.String(10), nullable=True)
    tipo_analisis = db.Column(db.String(100), nullable=True)
    microfonos = db.relationship('Microfono', backref='escenario', lazy=True)
    microfonos_historial = db.Column(db.Text, nullable=True)

class Microfono(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    identificador = db.Column(db.String(100), unique=True, nullable=False)
    modelo = db.Column(db.String(100), nullable=True)
    ubicacion = db.Column(db.String(100), nullable=True)
    escenario_id = db.Column(db.Integer, db.ForeignKey('escenario.id'), nullable=True)

class AudioResultado(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    audio_id = db.Column(db.Integer, db.ForeignKey('audio.id'), nullable=False)
    microfono_id = db.Column(db.Integer, db.ForeignKey('microfono.id'), nullable=True)
    escenario_id = db.Column(db.Integer, db.ForeignKey('escenario.id'), nullable=True)
    timestamp = db.Column(db.DateTime, default=lima_now)
    global_result = db.Column(db.Text)
    detailed_results = db.Column(db.Text)

with app.app_context():
    db.create_all()
    logger.debug("Base de datos y tablas creadas")


# Configuración de APScheduler para finalizar escenarios automáticamente
scheduler = BackgroundScheduler()
scheduler.start()

def finalizar_escenario_job(escenario_id):
    with app.app_context():
        escenario = Escenario.query.get(escenario_id)
        if escenario and escenario.estado == 'activo':
            escenario.estado = 'culminado'
            # Liberar micrófonos asignados
            for microfono in escenario.microfonos:
                microfono.escenario_id = None
            db.session.commit()
            logger.debug(f'Escenario {escenario_id} finalizado automáticamente.')

import soundfile as sf  # Añade esta importación al inicio del archivo

def analizar_audio_file(filepath, audio_id=None):
    """
    Analiza el archivo de audio y calcula tanto un resumen global como
    resultados detallados por intervalos de 1 segundo.
    Corrige encabezados WAV inválidos sin modificar el contenido original.
    """
    try:
        if audio_id:
            processing_audios[audio_id] = {
                "audio_id": audio_id,
                "start_time": datetime.datetime.utcnow().isoformat(),
                "duration": 0,
                "progress": 0,
                "finished": False,
                "status": "processing"
            }

        # 1. Leer el archivo con soundfile (tolera encabezados malformados)
        try:
            data, rate = sf.read(filepath, dtype='float32')
            data = data.astype(np.float64)  # Convertir al formato esperado
        except Exception as e:
            # 2. Si falla, leer con scipy y reparar datos
            rate, data = wavfile.read(filepath)
            if len(data.shape) > 1:
                data = data[:, 0]
            if data.dtype == np.int16:
                data = data / 32768.0
            elif data.dtype == np.int32:
                data = data / 2147483648.0
            elif data.dtype == np.uint8:
                data = (data - 128) / 128.0
            data = data.astype(np.float64)

        T = len(data) / rate

        mean_squared = np.mean(data**2)
        Lp_eqT = 10 * np.log10(mean_squared / (p0**2)) if mean_squared > 0 else 0
        p_max = np.max(np.abs(data))
        Lp_max = 10 * np.log10((p_max**2) / (p0**2)) if p_max > 0 else 0
        ET = np.sum(data**2) / rate
        LE = 10 * np.log10(ET / E0) if ET > 0 else 0
        J_energy = ET
        LJ = 10 * np.log10(J_energy / J0) if J_energy > 0 else 0
        IT = mean_squared / (rho * c)
        LI = 10 * np.log10(IT / I0) if IT > 0 else 0

        global_result = {
            "Lp_eqT": round(Lp_eqT, 2),
            "Lp_max": round(Lp_max, 2),
            "ET": round(ET, 6),
            "LE": round(LE, 2),
            "J_energy": round(J_energy, 6),
            "LJ": round(LJ, 2),
            "IT": round(IT, 6),
            "LI": round(LI, 2),
            "duration": round(T, 2),
            "sample_rate": rate
        }
        # Si se excede el umbral seguro, se agrega una alerta
        if global_result["Lp_max"] > SAFE_NOISE_LEVEL:
            global_result["alert"] = f"Nivel de ruido seguro excedido: {global_result['Lp_max']} dB"

        # Procesamiento detallado por intervalos de 1 segundo
        chunk_size = 1.0
        num_chunks = int(np.floor(T / chunk_size))

        if audio_id:
            processing_audios[audio_id] = {
                "audio_id": audio_id,
                "start_time": datetime.datetime.utcnow().isoformat(),
                "duration": round(T, 2),
                "progress": 0,
                "finished": False,
                "status": "processing"
            }

        detailed_results = []
        for i in range(num_chunks):
            start_index = int(i * chunk_size * rate)
            end_index = int((i + 1) * chunk_size * rate)
            chunk = data[start_index:end_index]
            if len(chunk) == 0:
                continue
            mean_squared_chunk = np.mean(chunk**2)
            Lp_eqT_chunk = 10 * np.log10(mean_squared_chunk / (p0**2)) if mean_squared_chunk > 0 else 0
            p_max_chunk = np.max(np.abs(chunk))
            Lp_max_chunk = 10 * np.log10((p_max_chunk**2) / (p0**2)) if p_max_chunk > 0 else 0
            ET_chunk = np.sum(chunk**2) / rate
            LE_chunk = 10 * np.log10(ET_chunk / E0) if ET_chunk > 0 else 0
            J_energy_chunk = ET_chunk
            LJ_chunk = 10 * np.log10(J_energy_chunk / J0) if J_energy_chunk > 0 else 0
            IT_chunk = mean_squared_chunk / (rho * c)
            LI_chunk = 10 * np.log10(IT_chunk / I0) if IT_chunk > 0 else 0

            chunk_result = {
                "timestamp": round(i * chunk_size, 2),
                "Lp_eqT": round(Lp_eqT_chunk, 2),
                "Lp_max": round(Lp_max_chunk, 2),
                "ET": round(ET_chunk, 6),
                "LE": round(LE_chunk, 2),
                "J_energy": round(J_energy_chunk, 6),
                "LJ": round(LJ_chunk, 2),
                "IT": round(IT_chunk, 6),
                "LI": round(LI_chunk, 2)
            }
            # Generar alerta en el intervalo si se excede el umbral
            if Lp_max_chunk > SAFE_NOISE_LEVEL:
                chunk_result["alert"] = f"Alerta en segundo {i}: {round(Lp_max_chunk,2)} dB"
            detailed_results.append(chunk_result)
            if audio_id:
                progress = int((i + 1) * 100 / num_chunks)
                processing_audios[audio_id]["progress"] = progress

        if audio_id:
            processing_audios[audio_id]["progress"] = 100
            processing_audios[audio_id]["finished"] = True
            processing_audios[audio_id]["status"] = "finished"
            socketio.emit('audio_procesado', {'audio_id': audio_id}, namespace="/")

        return global_result, detailed_results

    except Exception as e:
        if audio_id:
            processing_audios[audio_id]["finished"] = True
            processing_audios[audio_id]["status"] = "failed"
        logger.error(f"Error en análisis: {e}")
        raise

@app.route('/descargar-app')
def descargar_app():
    return send_file('app.py', as_attachment=True)


# Rutas básicas y de archivos
@app.route('/comparacion.html')
def comparacion():
    return render_template("comparacion.html")

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    logger.debug(f"Solicitando archivo: {filename}")
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        title = request.form['title']
        source = request.form['source']
        audio_file = request.files['audio']
        if audio_file:
            filename = audio_file.filename
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            audio_file.save(filepath)
            new_audio = Audio(filename=filename, title=title, source=source)
            db.session.add(new_audio)
            db.session.commit()
            return redirect(url_for('index'))
    audios = Audio.query.all()
    return render_template('index.html', audios=audios)

@app.route('/analyze/<int:audio_id>', methods=['GET'])
def analyze_audio(audio_id):
    audio = Audio.query.get_or_404(audio_id)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], audio.filename)
    try:
        global_result, detailed_results = analizar_audio_file(filepath, audio_id=audio.id)
        return render_template('analysis.html', audio=audio, global_result=global_result, detailed_results=detailed_results)
    except Exception as e:
        logger.error(f"Error al analizar el archivo: {e}")
        return render_template('error.html', message="Error al analizar el archivo de audio.")

import traceback
import sys

@app.route('/upload_audio', methods=['POST'])
def upload_audio():
    try:
        logger.debug("=== INICIO DE SOLICITUD ===")
        logger.debug(f"Headers: {dict(request.headers)}")
        logger.debug(f"Form data: {request.form}")
        logger.debug(f"Archivos: {request.files}")

        # Verificar API Key
        api_key = request.headers.get('X-API-Key')
        if not api_key:
            logger.error("API Key no proporcionada")
            return jsonify({"error": "API key requerida"}), 401
        if api_key != app.config['API_KEY']:
            logger.error("API Key inválida")
            return jsonify({"error": "API key inválida"}), 403

        # Verificar campo 'audio'
        if 'audio' not in request.files:
            logger.error("Campo 'audio' no encontrado")
            return jsonify({"error": "Campo 'audio' requerido"}), 400

        audio_file = request.files['audio']
        if audio_file.filename == '':
            logger.error("Nombre de archivo vacío")
            return jsonify({"error": "Nombre de archivo inválido"}), 400

        # Sanitizar el nombre del archivo y construir ruta ABSOLUTA
        from werkzeug.utils import secure_filename  # Añade esta línea al inicio del archivo
        filename = secure_filename(audio_file.filename)
        upload_folder = os.path.abspath(app.config['UPLOAD_FOLDER'])  # Ruta absoluta
        filepath = os.path.join(upload_folder, filename)

        # Asegurar que el directorio existe
        os.makedirs(upload_folder, exist_ok=True)  # Crea la carpeta si no existe

        # Guardar el archivo con logs de depuración
        logger.debug(f"Ruta ABSOLUTA de guardado: {filepath}")
        audio_file.save(filepath)
        logger.debug(f"Archivo guardado exitosamente en: {filepath}")

        # Registrar audio en BD
        new_audio = Audio(
            filename=filename,
            title=request.form.get('title', 'Audio desde Arduino'),
            source=request.form.get('source', 'Desconocido')
        )
        db.session.add(new_audio)
        db.session.commit()
        logger.debug(f"Audio registrado en BD con ID {new_audio.id}")

        # Notificar con Socket.IO
        socketio.emit('new_audio', {
            'audio_id': new_audio.id,
            'message': 'Nuevo audio cargado'
        }, namespace="/", to="")

        logger.debug(f"Evento SocketIO 'new_audio' emitido para audio ID {new_audio.id}")

        # Procesar el audio
        global_result, detailed_results = analizar_audio_file(filepath, audio_id=new_audio.id)

        # Asignación de micrófono
        microfono_id = request.form.get('microfono_id')
        indicador_asignacion = "No asignado"
        if microfono_id:
            microfono = Microfono.query.get(microfono_id)
            if microfono and microfono.escenario_id is not None:
                resultado = AudioResultado(
                    audio_id=new_audio.id,
                    microfono_id=microfono.id,
                    escenario_id=microfono.escenario_id,
                    global_result=json.dumps(global_result),
                    detailed_results=json.dumps(detailed_results)
                )
                db.session.add(resultado)
                db.session.commit()
                indicador_asignacion = "Asignado a escenario"
                logger.debug(f"Resultados almacenados para audio ID {new_audio.id} (Micrófono {microfono.id}, Escenario {microfono.escenario_id})")
            else:
                logger.debug("El micrófono indicado no está asignado a ningún escenario o no existe")
        else:
            logger.debug("No se proporcionó ID de micrófono en la solicitud")

        response = {
            "message": "Audio subido y analizado",
            "filename": filename,
            "size": os.path.getsize(filepath),
            "audio_id": new_audio.id,
            "asignacion": indicador_asignacion,
         #   "global_result": global_result,
          #  "detailed_results": detailed_results
        }
        return jsonify(response), 200

    except Exception as e:
        error_type, error_value, error_traceback = sys.exc_info()
        trace = "".join(traceback.format_tb(error_traceback))

        logger.exception(f"Error durante la subida y análisis del audio:\n{trace}")

        error_details = {
            "error": str(e),
            "error_type": str(error_type),
            "trace": trace
        }

        # Si el error viene de Nginx o Gunicorn, agregar info
        if "nginx" in str(error_type).lower() or "gunicorn" in str(error_type).lower():
            error_details["server_issue"] = "El error puede estar relacionado con Nginx/Gunicorn"

        return jsonify(error_details), 500

    finally:
        logger.debug("=== FIN DE SOLICITUD ===")
# Endpoints para escenarios y micrófonos

@app.route('/escenarios', methods=['POST'])
def crear_escenario():
    data = request.json
    try:
        nombre = data['nombre']
        descripcion = data.get('descripcion', '')
        ubicacion = data.get('ubicacion', '')
        start_time_str = data['start_time']
        end_time_str = data['end_time']
        start_time = datetime.datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
        end_time = datetime.datetime.strptime(end_time_str, "%Y-%m-%d %H:%M:%S")
        if start_time >= end_time:
            return jsonify({'error': 'La fecha de inicio debe ser anterior a la fecha de fin'}), 400

        estado = 'programado'
        duracion_segundos = (end_time - start_time).total_seconds()
        horas_medicion = round(duracion_segundos / 3600.0, 2)
        dias_medicion = round(duracion_segundos / (3600.0 * 24), 2)

        tipo_ruido = data.get('tipo_ruido', '')
        num_fuentes = data.get('num_fuentes', 0)
        num_personas = data.get('num_personas', 0)
        proteccion_auditiva = data.get('proteccion_auditiva', '')
        tipo_analisis = data.get('tipo_analisis', '')

        nuevo_escenario = Escenario(
            nombre=nombre,
            descripcion=descripcion,
            ubicacion=ubicacion,
            start_time=start_time,
            end_time=end_time,
            estado=estado,
            horas_medicion=horas_medicion,
            dias_medicion=dias_medicion,
            tipo_ruido=tipo_ruido,
            num_fuentes=num_fuentes,
            num_personas=num_personas,
            proteccion_auditiva=proteccion_auditiva,
            tipo_analisis=tipo_analisis
        )
        db.session.add(nuevo_escenario)
        db.session.commit()
        return jsonify({'mensaje': 'Escenario creado', 'id': nuevo_escenario.id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/escenarios', methods=['GET'])
def listar_escenarios():
    escenarios = Escenario.query.all()
    resultado = []
    for e in escenarios:
        if e.estado == 'culminado' and e.microfonos_historial:
            try:
                mic_ids = json.loads(e.microfonos_historial)
                mic_list = []
                for mid in mic_ids:
                    mic = Microfono.query.get(mid)
                    if mic:
                        mic_list.append({
                            'id': mic.id,
                            'identificador': mic.identificador,
                            'modelo': mic.modelo,
                            'ubicacion': mic.ubicacion
                        })
                microfonos = mic_list
            except Exception as ex:
                microfonos = []
        else:
            microfonos = [{
                'id': m.id,
                'identificador': m.identificador,
                'modelo': m.modelo,
                'ubicacion': m.ubicacion
            } for m in e.microfonos]
        resultado.append({
            'id': e.id,
            'nombre': e.nombre,
            'descripcion': e.descripcion,
            'ubicacion': e.ubicacion,
            'start_time': e.start_time.strftime("%Y-%m-%d %H:%M:%S"),
            'end_time': e.end_time.strftime("%Y-%m-%d %H:%M:%S"),
            'estado': e.estado,
            'horas_medicion': e.horas_medicion,
            'dias_medicion': e.dias_medicion,
            'tipo_ruido': e.tipo_ruido,
            'num_fuentes': e.num_fuentes,
            'num_personas': e.num_personas,
            'proteccion_auditiva': e.proteccion_auditiva,
            'tipo_analisis': e.tipo_analisis,
            'microfonos': microfonos
        })
    return jsonify(resultado)

@app.route('/micros', methods=['GET'])
def listar_microfonos():
    microfonos = Microfono.query.all()
    return jsonify([{
        'id': m.id,
        'identificador': m.identificador,
        'modelo': m.modelo,
        'ubicacion': m.ubicacion,
        'escenario_id': m.escenario_id,
        'escenario_nombre': m.escenario.nombre if m.escenario else "Libre"
    } for m in microfonos])

def _actualizar_campos_descriptivos(escenario, data):
    """Update descriptive metadata fields shared between culminado and programado scenarios."""
    escenario.nombre = data.get('nombre', escenario.nombre)
    escenario.descripcion = data.get('descripcion', escenario.descripcion)
    escenario.ubicacion = data.get('ubicacion', escenario.ubicacion)
    escenario.tipo_ruido = data.get('tipo_ruido', escenario.tipo_ruido)
    escenario.num_fuentes = data.get('num_fuentes', escenario.num_fuentes)
    escenario.num_personas = data.get('num_personas', escenario.num_personas)
    escenario.proteccion_auditiva = data.get('proteccion_auditiva', escenario.proteccion_auditiva)
    escenario.tipo_analisis = data.get('tipo_analisis', escenario.tipo_analisis)


@app.route('/escenarios/<int:id>', methods=['PUT'])
def actualizar_escenario(id):
    data = request.json
    escenario = Escenario.query.get_or_404(id)
    try:
        # For completed scenarios, allow editing descriptive metadata only (no date changes)
        if escenario.estado == 'culminado':
            _actualizar_campos_descriptivos(escenario, data)
            db.session.commit()
            return jsonify({'mensaje': 'Datos del escenario actualizados'})
        if escenario.estado != 'programado':
            return jsonify({'error': 'Solo se pueden modificar escenarios programados'}), 400
        _actualizar_campos_descriptivos(escenario, data)
        if 'start_time' in data:
            escenario.start_time = datetime.datetime.strptime(data['start_time'], "%Y-%m-%d %H:%M:%S")
        if 'end_time' in data:
            escenario.end_time = datetime.datetime.strptime(data['end_time'], "%Y-%m-%d %H:%M:%S")
        duracion_segundos = (escenario.end_time - escenario.start_time).total_seconds()
        escenario.horas_medicion = round(duracion_segundos / 3600.0, 2)
        escenario.dias_medicion = round(duracion_segundos / (3600.0 * 24), 2)
        db.session.commit()
        return jsonify({'mensaje': 'Escenario actualizado'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/escenarios/<int:id>', methods=['DELETE'])
def eliminar_escenario(id):
    escenario = Escenario.query.get_or_404(id)
    if escenario.estado != 'programado':
        return jsonify({'error': 'No se puede eliminar un escenario iniciado o finalizado'}), 400
    db.session.delete(escenario)
    db.session.commit()
    return jsonify({'mensaje': 'Escenario eliminado'})

@app.route('/escenarios/<int:escenario_id>/iniciar', methods=['POST'])
def iniciar_escenario(escenario_id):
    escenario = Escenario.query.get_or_404(escenario_id)
    if escenario.estado != 'programado':
        return jsonify({'error': 'El escenario ya se ha iniciado o finalizado'}), 400
    escenario.estado = 'activo'
    db.session.commit()
    now = datetime.datetime.now()
    if escenario.end_time > now:
        scheduler.add_job(func=finalizar_escenario_job, trigger='date', run_date=escenario.end_time, args=[escenario.id])
    else:
        finalizar_escenario_job(escenario.id)
    return jsonify({'mensaje': 'Escenario iniciado', 'finalizacion_programada': escenario.end_time.isoformat()})

@app.route('/escenarios/<int:escenario_id>/asignar_microfonos', methods=['POST'])
def asignar_microfonos(escenario_id):
    escenario = Escenario.query.get_or_404(escenario_id)
    if escenario.estado != 'programado':
        return jsonify({'error': 'No se pueden asignar micrófonos a un escenario iniciado o finalizado'}), 400
    data = request.json
    microfonos = Microfono.query.filter(Microfono.id.in_(data['microfono_ids']), Microfono.escenario_id.is_(None)).all()
    if len(microfonos) != len(data['microfono_ids']):
        return jsonify({'error': 'Algunos micrófonos ya están en uso'}), 400
    for microfono in microfonos:
        microfono.escenario_id = escenario.id
    db.session.commit()
    return jsonify({'mensaje': 'Micrófonos asignados'})

@app.route('/micros', methods=['POST'])
def crear_microfono():
    data = request.json
    try:
        identificador = data['identificador']
        modelo = data.get('modelo', '')
        ubicacion = data.get('ubicacion', '')
        nuevo_microfono = Microfono(identificador=identificador, modelo=modelo, ubicacion=ubicacion)
        db.session.add(nuevo_microfono)
        db.session.commit()
        return jsonify({'mensaje': 'Micrófono creado', 'id': nuevo_microfono.id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/micros/<int:id>', methods=['DELETE'])
def eliminar_microfono(id):
    microfono = Microfono.query.get_or_404(id)
    if microfono.escenario_id is not None:
        return jsonify({'error': 'No se puede eliminar un micrófono asignado a un escenario activo'}), 400
    db.session.delete(microfono)
    db.session.commit()
    return jsonify({'mensaje': 'Micrófono eliminado'})

@app.route('/escenarios/<int:escenario_id>/finalizar', methods=['POST'])
def finalizar_escenario_manual(escenario_id):
    escenario = Escenario.query.get_or_404(escenario_id)
    if escenario.estado != 'activo':
        return jsonify({'error': 'El escenario no está en ejecución'}), 400
    mic_list = list(escenario.microfonos)
    historial = [m.id for m in mic_list]
    escenario.microfonos_historial = json.dumps(historial)
    escenario.estado = 'culminado'
    for microfono in mic_list:
        microfono.escenario_id = None
    db.session.commit()
    return jsonify({'mensaje': 'Escenario finalizado y se conserva el historial de micrófonos'})

@app.route('/escenarios/<int:escenario_id>/detalle', methods=['GET'])
def detalle_escenario(escenario_id):
    escenario = Escenario.query.get_or_404(escenario_id)
    resultados = AudioResultado.query.filter_by(escenario_id=escenario_id).order_by(AudioResultado.timestamp.desc()).all()
    microfonos_resultado = {}
    for res in resultados:
        if res.microfono_id not in microfonos_resultado:
            microfonos_resultado[res.microfono_id] = res
    microfonos_data = []
    for micro_id, resultado in microfonos_resultado.items():
        micro = Microfono.query.get(micro_id)
        if micro:
            micro_data = {
                'id': micro.id,
                'identificador': micro.identificador,
                'resumen': json.loads(resultado.global_result) if resultado.global_result else {},
                'detalle': json.loads(resultado.detailed_results) if resultado.detailed_results else {}
            }
            microfonos_data.append(micro_data)
    data = {
        'id': escenario.id,
        'nombre': escenario.nombre,
        'descripcion': escenario.descripcion,
        'ubicacion': escenario.ubicacion,
        'start_time': escenario.start_time.strftime("%Y-%m-%d %H:%M:%S"),
        'end_time': escenario.end_time.strftime("%Y-%m-%d %H:%M:%S"),
        'estado': escenario.estado,
        'horas_medicion': escenario.horas_medicion,
        'dias_medicion': escenario.dias_medicion,
        'tipo_ruido': escenario.tipo_ruido,
        'num_fuentes': escenario.num_fuentes,
        'num_personas': escenario.num_personas,
        'proteccion_auditiva': escenario.proteccion_auditiva,
        'tipo_analisis': escenario.tipo_analisis,
        'microfonos': microfonos_data
    }
    return jsonify(data)

@app.route('/micros/<int:microfono_id>/resultado', methods=['GET'])
def detalle_microfono(microfono_id):
    resultado = AudioResultado.query.filter_by(microfono_id=microfono_id).order_by(AudioResultado.timestamp.desc()).first()
    if resultado:
        data = {
            'audio_id': resultado.audio_id,
            'microfono_id': resultado.microfono_id,
            'escenario_id': resultado.escenario_id,
            'timestamp': resultado.timestamp.isoformat(),
            'global_result': json.loads(resultado.global_result) if resultado.global_result else {},
            'detailed_results': json.loads(resultado.detailed_results) if resultado.detailed_results else {}
        }
        return jsonify(data)
    else:
        return jsonify({'error': 'No hay resultados para este micrófono'}), 404

@app.route('/audios_sin_asignacion', methods=['GET'])
def audios_sin_asignacion():
    subquery = db.session.query(AudioResultado.audio_id)
    audios = Audio.query.filter(~Audio.id.in_(subquery)).all()
    result = [{'id': a.id, 'filename': a.filename, 'title': a.title, 'source': a.source} for a in audios]
    return jsonify(result)

@app.route('/escenarios/<int:escenario_id>/microfono/<int:microfono_id>/audios', methods=['GET'])
def obtener_audios(escenario_id, microfono_id):
    resultados = AudioResultado.query.filter_by(
        escenario_id=escenario_id,
        microfono_id=microfono_id
    ).order_by(AudioResultado.timestamp.asc()).all()
    audios = []
    for res in resultados:
        try:
            global_result = json.loads(res.global_result) if res.global_result else {}
            detailed_results = json.loads(res.detailed_results) if res.detailed_results else {}
        except Exception as e:
            global_result = {}
            detailed_results = {}
        audio_entry = Audio.query.get(res.audio_id)
        filename = audio_entry.filename if audio_entry else ""
        audios.append({
            "audio_id": res.audio_id,
            "timestamp": res.timestamp.isoformat(),
            "global_result": global_result,
            "detailed_results": detailed_results,
            "filename": filename
        })
    return jsonify(audios)

# Modificaciones necesarias en app.py

# REEMPLAZAR la ruta existente /monitoring/<int:escenario_id>/microfono/<int:microfono_id>
@app.route('/monitoring/<int:escenario_id>/microfono/<int:microfono_id>')
def monitoring(escenario_id, microfono_id):
    escenario = Escenario.query.get_or_404(escenario_id)
    
    # NUEVA LÓGICA: Verificar el estado del escenario
    if escenario.estado == 'culminado':
        # Si el escenario está culminado, renderizar la nueva plantilla
        return render_template(
            'monitoring_finished.html',  # NUEVO TEMPLATE
            escenario_id=escenario_id,
            microfono_id=microfono_id,
            escenario_start=escenario.start_time.isoformat(),
            escenario_end=escenario.end_time.isoformat(),
            limite_referencia=85.0  # Línea de referencia en 85 dB
        )
    else:
        # Si está activo, usar el template original
        return render_template(
            'monitoring.html',
            escenario_id=escenario_id,
            microfono_id=microfono_id,
            escenario_start=escenario.start_time.isoformat(),
            escenario_end=escenario.end_time.isoformat()
        )

# AGREGAR nuevo endpoint para estadísticas de excesos del límite
@app.route('/escenarios/<int:escenario_id>/microfono/<int:mic_id>/analisis_excesos', methods=['GET'])
def analisis_excesos(escenario_id, mic_id):
    """
    Análisis de cuántas veces se excedió el límite de 85 dB en el escenario
    """
    resultados = AudioResultado.query.filter_by(escenario_id=escenario_id, microfono_id=mic_id).all()
    if not resultados:
        return jsonify({})
    
    limite_seguro = 85.0
    total_puntos = 0
    puntos_excedidos = 0
    excesos_por_audio = []
    max_exceso = 0
    duracion_total_exceso = 0
    
    for res in resultados:
        try:
            global_result = json.loads(res.global_result)
            detailed_results = json.loads(res.detailed_results)
            
            audio_excesos = 0
            duracion_audio_exceso = 0
            
            # Analizar cada punto detallado del audio
            for punto in detailed_results:
                total_puntos += 1
                lp_max = float(punto.get('Lp_max', 0))
                
                if lp_max > limite_seguro:
                    puntos_excedidos += 1
                    audio_excesos += 1
                    duracion_audio_exceso += 1  # Cada punto representa ~1 segundo
                    
                    if lp_max > max_exceso:
                        max_exceso = lp_max
            
            duracion_total_exceso += duracion_audio_exceso
            
            excesos_por_audio.append({
                "audio_id": res.audio_id,
                "timestamp": res.timestamp.isoformat(),
                "puntos_excedidos": audio_excesos,
                "duracion_exceso_segundos": duracion_audio_exceso,
                "nivel_maximo": float(global_result.get("Lp_max", 0))
            })
            
        except Exception as e:
            continue
    
    porcentaje_exceso = (puntos_excedidos / total_puntos * 100) if total_puntos > 0 else 0
    
    analysis = {
        "limite_referencia": limite_seguro,
        "total_puntos_medidos": total_puntos,
        "puntos_que_excedieron": puntos_excedidos,
        "porcentaje_exceso": round(porcentaje_exceso, 2),
        "duracion_total_exceso_segundos": duracion_total_exceso,
        "duracion_total_exceso_minutos": round(duracion_total_exceso / 60, 2),
        "nivel_maximo_registrado": round(max_exceso, 2),
        "cantidad_audios": len(resultados),
        "excesos_por_audio": excesos_por_audio,
        "estado": "ANÁLISIS COMPLETADO",
        "evaluacion_seguridad": "SEGURO" if puntos_excedidos == 0 else "ATENCIÓN REQUERIDA"
    }
    return jsonify(analysis)


@app.route('/escenarios/<int:escenario_id>/microfono/<int:mic_id>/analisis_general', methods=['GET'])
def analisis_general(escenario_id, mic_id):
    # Recupera todos los resultados para ese escenario y micrófono
    resultados = AudioResultado.query.filter_by(escenario_id=escenario_id, microfono_id=mic_id).all()
    if not resultados:
        return jsonify({})
    
    # Inicializamos acumuladores (en escala lineal) para combinar los audios
    total_duration = 0.0
    total_ET = 0.0
    weighted_sum_pressure_sq = 0.0
    for res in resultados:
        global_result = json.loads(res.global_result)
        # Asegurarse de convertir a número
        duration = float(global_result.get("duration", 0))
        Lp_eqT = float(global_result.get("Lp_eqT", 0))
        ET = float(global_result.get("ET", 0))
        # Convertir Lp_eqT (dB) a p² en escala lineal:  
        #   p² = p₀² · 10^(Lp_eqT/10)
        pressure_sq = (2.0e-5)**2 * (10 ** (Lp_eqT / 10)) if Lp_eqT else 0

        total_duration += duration
        total_ET += ET
        weighted_sum_pressure_sq += pressure_sq * duration
    
    mean_pressure_sq = weighted_sum_pressure_sq / total_duration if total_duration > 0 else 0
    # Calcular Lp_eqT global a partir del promedio ponderado en escala lineal
    Lp_eqT_global = 10 * np.log10(mean_pressure_sq / (2.0e-5)**2) if mean_pressure_sq > 0 else 0
    # Calcular el nivel de exposición acústica global
    LE_global = 10 * np.log10(total_ET / 4.0e-10) if total_ET > 0 else 0
    # Calcular L_EX,8h según NTP ISO 9612:2010 (exposición normalizada a 8 horas)
    duracion_horas = total_duration / 3600.0
    # Require at least 1 minute of measurement for a meaningful L_EX,8h
    L_EX_8h = round(Lp_eqT_global + 10 * np.log10(duracion_horas / 8.0), 2) if duracion_horas >= (1.0 / 60.0) else None

    analysis = {
         "duration": total_duration,
         "ET": total_ET,
         "Lp_eqT": round(Lp_eqT_global, 2),
         "LE": round(LE_global, 2),
         "L_EX_8h": L_EX_8h,
         "cantidad_audios": len(resultados)
    }
    return jsonify(analysis)


# NUEVOS ENDPOINTS PARA EL DASHBOARD

@app.route('/audio_status/<int:audio_id>')
def audio_status(audio_id):
    """
    Devuelve el estado de procesamiento de un audio (progreso, tiempo de inicio y si terminó).
    """
    status = processing_audios.get(audio_id)
    if status:
        return jsonify(status)
    else:
        return jsonify({"error": "No se encontró el audio o ya finalizó"}), 404

@app.route('/dashboard_data')
def dashboard_data():
    """
    Reúne información global para el dashboard:
      - Total de audios almacenados.
      - Total de micrófonos registrados.
      - Total de micrófonos activos (aquellos asignados a un escenario).
      - Los 10 audios más recientes, con el tiempo transcurrido, micrófono y escenario asociados.
      - Alertas generadas durante el análisis.
      - Datos de audios en procesamiento.
      - Datos para el gráfico de niveles de ruido (noise_data).
      - Frecuencia de llegada de audios por micrófono activo (frequency_data) por hora.
    """
    now = datetime.datetime.now()
    total_audios = Audio.query.count()
    total_microphones = Microfono.query.count()
    active_microphones = Microfono.query.filter(Microfono.escenario_id.isnot(None)).count()

    recent_results = AudioResultado.query.order_by(AudioResultado.timestamp.desc()).limit(10).all()
    recent_audios = []
    alerts = []
    for res in recent_results:
        audio = Audio.query.get(res.audio_id)
        mic = Microfono.query.get(res.microfono_id) if res.microfono_id else None
        escenario = Escenario.query.get(res.escenario_id) if res.escenario_id else None
        time_since = (now - res.timestamp).total_seconds()  # en segundos
        global_result = json.loads(res.global_result) if res.global_result else {}
        if "alert" in global_result:
            alerts.append({
                "audio_id": res.audio_id,
                "mic": mic.identificador if mic else "N/A",
                "escenario": escenario.nombre if escenario else "N/A",
                "alert": global_result["alert"]
            })
        recent_audios.append({
            "audio_id": res.audio_id,
            "timestamp": res.timestamp.isoformat(),
            "time_since": time_since,
            "mic": mic.identificador if mic else "N/A",
            "escenario": escenario.nombre if escenario else "N/A",
            "global_result": global_result
        })

    # Generar datos reales para el gráfico de niveles de ruido
    noise_data = {"labels": [], "levels": []}
    if recent_results:
        first_result = recent_results[0]
        detailed = json.loads(first_result.detailed_results) if first_result.detailed_results else []
        for d in detailed:
            noise_data["labels"].append(f"{d.get('timestamp', 0)} s")
            noise_data["levels"].append(d.get("Lp_max", 0))
    else:
        noise_data = {"labels": ["0 s", "1 s", "2 s", "3 s", "4 s", "5 s", "6 s", "7 s", "8 s", "9 s"],
                      "levels": [0] * 10}

    # Calcular la frecuencia de llegada de audios por micrófono activo en la última hora
    one_hour_ago = now - datetime.timedelta(hours=1)
    frequency_data = []
    active_mics = Microfono.query.filter(Microfono.escenario_id.isnot(None)).all()
    for mic in active_mics:
        count = AudioResultado.query.filter(
            AudioResultado.microfono_id == mic.id,
            AudioResultado.timestamp >= one_hour_ago
        ).count()
        frequency_data.append({"identificador": mic.identificador, "frequency": count})

    # Registrar en el log del servidor la información obtenida
    app.logger.info("Dashboard Data: total_audios=%s, total_microphones=%s, active_microphones=%s",
                    total_audios, total_microphones, active_microphones),
    app.logger.info("Alerts: %s", alerts)


    return jsonify({
        "total_audios": total_audios,
        "total_microphones": total_microphones,
        "active_microphones": active_microphones,
        "recent_audios": recent_audios,
        "alerts": alerts,
        "processing": list(processing_audios.values()),
        "noise_data": noise_data,
        "frequency_data": frequency_data
    })


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
