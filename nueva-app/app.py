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
import soundfile as sf
import traceback
import sys
from werkzeug.utils import secure_filename

# Configurar logger
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
CORS(app, origins=["http://localhost:3000", "https://your-nextjs-domain.com"])  # Add your Next.js domain
socketio = SocketIO(app, cors_allowed_origins=["http://localhost:3000", "https://your-nextjs-domain.com"])
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
            for microfono in escenario.microfonos:
                microfono.escenario_id = None
            db.session.commit()
            logger.debug(f'Escenario {escenario_id} finalizado automáticamente.')

def analizar_audio_file(filepath, audio_id=None):
    """
    Analiza el archivo de audio y calcula tanto un resumen global como
    resultados detallados por intervalos de 1 segundo.
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

        # Leer el archivo con soundfile
        try:
            data, rate = sf.read(filepath, dtype='float32')
            data = data.astype(np.float64)
        except Exception as e:
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

# ============================================================================
# API ROUTES FOR NEXT.JS FRONTEND
# ============================================================================

# Health check endpoint
@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.datetime.utcnow().isoformat(),
        "version": "1.0.0"
    })

# Get all audios with pagination
@app.route('/api/audios', methods=['GET'])
def get_audios():
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 10, type=int)
    
    audios = Audio.query.order_by(Audio.timestamp.desc()).paginate(
        page=page, per_page=per_page, error_out=False
    )
    
    return jsonify({
        "audios": [{
            "id": audio.id,
            "filename": audio.filename,
            "title": audio.title,
            "source": audio.source,
            "timestamp": audio.timestamp.isoformat()
        } for audio in audios.items],
        "total": audios.total,
        "pages": audios.pages,
        "current_page": page,
        "per_page": per_page
    })

# Get single audio details
@app.route('/api/audios/<int:audio_id>', methods=['GET'])
def get_audio(audio_id):
    audio = Audio.query.get_or_404(audio_id)
    return jsonify({
        "id": audio.id,
        "filename": audio.filename,
        "title": audio.title,
        "source": audio.source,
        "timestamp": audio.timestamp.isoformat()
    })

# Delete audio
@app.route('/api/audios/<int:audio_id>', methods=['DELETE'])
def delete_audio(audio_id):
    audio = Audio.query.get_or_404(audio_id)
    
    # Delete associated results
    AudioResultado.query.filter_by(audio_id=audio_id).delete()
    
    # Delete file from filesystem
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], audio.filename)
    if os.path.exists(filepath):
        os.remove(filepath)
    
    db.session.delete(audio)
    db.session.commit()
    
    return jsonify({"message": "Audio eliminado exitosamente"})

# Upload and analyze audio (enhanced for Next.js)
@app.route('/api/upload', methods=['POST'])
def upload_audio_api():
    try:
        logger.debug("=== INICIO DE SOLICITUD API ===")
        logger.debug(f"Headers: {dict(request.headers)}")
        logger.debug(f"Form data: {request.form}")
        logger.debug(f"Archivos: {request.files}")

        # Verificar API Key (opcional para frontend)
        api_key = request.headers.get('X-API-Key')
        if api_key and api_key != app.config['API_KEY']:
            return jsonify({"error": "API key inválida"}), 403

        if 'audio' not in request.files:
            return jsonify({"error": "Campo 'audio' requerido"}), 400

        audio_file = request.files['audio']
        if audio_file.filename == '':
            return jsonify({"error": "Nombre de archivo inválido"}), 400

        filename = secure_filename(audio_file.filename)
        upload_folder = os.path.abspath(app.config['UPLOAD_FOLDER'])
        filepath = os.path.join(upload_folder, filename)

        os.makedirs(upload_folder, exist_ok=True)
        audio_file.save(filepath)

        new_audio = Audio(
            filename=filename,
            title=request.form.get('title', 'Audio desde Frontend'),
            source=request.form.get('source', 'Web Upload')
        )
        db.session.add(new_audio)
        db.session.commit()

        # Emit socket event
        socketio.emit('new_audio', {
            'audio_id': new_audio.id,
            'message': 'Nuevo audio cargado'
        }, namespace="/")

        # Process audio
        global_result, detailed_results = analizar_audio_file(filepath, audio_id=new_audio.id)

        # Handle microphone assignment
        microfono_id = request.form.get('microfono_id')
        asignacion_status = "No asignado"
        
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
                asignacion_status = "Asignado a escenario"

        response = {
            "success": True,
            "message": "Audio subido y analizado exitosamente",
            "data": {
                "audio_id": new_audio.id,
                "filename": filename,
                "size": os.path.getsize(filepath),
                "asignacion": asignacion_status,
                "global_result": global_result,
                "detailed_results": detailed_results[:10]  # First 10 seconds for preview
            }
        }
        return jsonify(response), 200

    except Exception as e:
        logger.exception("Error durante la subida y análisis del audio")
        return jsonify({
            "success": False,
            "error": str(e),
            "message": "Error al procesar el audio"
        }), 500

# Get audio analysis results
@app.route('/api/audios/<int:audio_id>/analysis', methods=['GET'])
def get_audio_analysis(audio_id):
    audio = Audio.query.get_or_404(audio_id)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], audio.filename)
    
    if not os.path.exists(filepath):
        return jsonify({"error": "Archivo de audio no encontrado"}), 404
    
    try:
        global_result, detailed_results = analizar_audio_file(filepath)
        return jsonify({
            "audio": {
                "id": audio.id,
                "filename": audio.filename,
                "title": audio.title,
                "source": audio.source,
                "timestamp": audio.timestamp.isoformat()
            },
            "global_result": global_result,
            "detailed_results": detailed_results
        })
    except Exception as e:
        return jsonify({"error": f"Error al analizar el archivo: {str(e)}"}), 500

# Enhanced scenarios endpoints
@app.route('/api/escenarios', methods=['GET'])
def get_escenarios():
    escenarios = Escenario.query.all()
    resultado = []
    for e in escenarios:
        # Handle microphones based on scenario state
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
            except:
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
    return jsonify({"escenarios": resultado})

@app.route('/api/escenarios', methods=['POST'])
def create_escenario():
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

        duracion_segundos = (end_time - start_time).total_seconds()
        horas_medicion = round(duracion_segundos / 3600.0, 2)
        dias_medicion = round(duracion_segundos / (3600.0 * 24), 2)

        nuevo_escenario = Escenario(
            nombre=nombre,
            descripcion=descripcion,
            ubicacion=ubicacion,
            start_time=start_time,
            end_time=end_time,
            estado='programado',
            horas_medicion=horas_medicion,
            dias_medicion=dias_medicion,
            tipo_ruido=data.get('tipo_ruido', ''),
            num_fuentes=data.get('num_fuentes', 0),
            num_personas=data.get('num_personas', 0),
            proteccion_auditiva=data.get('proteccion_auditiva', ''),
            tipo_analisis=data.get('tipo_analisis', '')
        )
        db.session.add(nuevo_escenario)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'Escenario creado exitosamente',
            'data': {'id': nuevo_escenario.id}
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# Enhanced microphones endpoints
@app.route('/api/microfonos', methods=['GET'])
def get_microfonos():
    microfonos = Microfono.query.all()
    return jsonify({
        "microfonos": [{
            'id': m.id,
            'identificador': m.identificador,
            'modelo': m.modelo,
            'ubicacion': m.ubicacion,
            'escenario_id': m.escenario_id,
            'escenario_nombre': m.escenario.nombre if m.escenario else "Libre",
            'estado': 'Asignado' if m.escenario_id else 'Libre'
        } for m in microfonos]
    })

@app.route('/api/microfonos', methods=['POST'])
def create_microfono():
    data = request.json
    try:
        nuevo_microfono = Microfono(
            identificador=data['identificador'],
            modelo=data.get('modelo', ''),
            ubicacion=data.get('ubicacion', '')
        )
        db.session.add(nuevo_microfono)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'Micrófono creado exitosamente',
            'data': {'id': nuevo_microfono.id}
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# Dashboard data endpoint (enhanced)
@app.route('/api/dashboard', methods=['GET'])
def get_dashboard_data():
    now = datetime.datetime.now()
    total_audios = Audio.query.count()
    total_microphones = Microfono.query.count()
    active_microphones = Microfono.query.filter(Microfono.escenario_id.isnot(None)).count()
    active_scenarios = Escenario.query.filter_by(estado='activo').count()

    # Recent audios with more details
    recent_results = AudioResultado.query.order_by(AudioResultado.timestamp.desc()).limit(10).all()
    recent_audios = []
    alerts = []
    
    for res in recent_results:
        audio = Audio.query.get(res.audio_id)
        mic = Microfono.query.get(res.microfono_id) if res.microfono_id else None
        escenario = Escenario.query.get(res.escenario_id) if res.escenario_id else None
        time_since = (now - res.timestamp).total_seconds()
        global_result = json.loads(res.global_result) if res.global_result else {}
        
        if "alert" in global_result:
            alerts.append({
                "audio_id": res.audio_id,
                "mic": mic.identificador if mic else "N/A",
                "escenario": escenario.nombre if escenario else "N/A",
                "alert": global_result["alert"],
                "timestamp": res.timestamp.isoformat()
            })
        
        recent_audios.append({
            "audio_id": res.audio_id,
            "timestamp": res.timestamp.isoformat(),
            "time_since": time_since,
            "mic": mic.identificador if mic else "N/A",
            "escenario": escenario.nombre if escenario else "N/A",
            "global_result": global_result
        })

    # Generate noise level data for charts
    noise_data = {"labels": [], "levels": []}
    if recent_results:
        first_result = recent_results[0]
        detailed = json.loads(first_result.detailed_results) if first_result.detailed_results else []
        for d in detailed[:20]:  # Limit to first 20 seconds
            noise_data["labels"].append(f"{d.get('timestamp', 0)} s")
            noise_data["levels"].append(d.get("Lp_max", 0))

    # Audio frequency data per microphone
    one_hour_ago = now - datetime.timedelta(hours=1)
    frequency_data = []
    active_mics = Microfono.query.filter(Microfono.escenario_id.isnot(None)).all()
    for mic in active_mics:
        count = AudioResultado.query.filter(
            AudioResultado.microfono_id == mic.id,
            AudioResultado.timestamp >= one_hour_ago
        ).count()
        frequency_data.append({
            "identificador": mic.identificador,
            "frequency": count,
            "escenario": mic.escenario.nombre if mic.escenario else "N/A"
        })

    return jsonify({
        "summary": {
            "total_audios": total_audios,
            "total_microphones": total_microphones,
            "active_microphones": active_microphones,
            "active_scenarios": active_scenarios
        },
        "recent_audios": recent_audios,
        "alerts": alerts,
        "processing": list(processing_audios.values()),
        "charts": {
            "noise_data": noise_data,
            "frequency_data": frequency_data
        },
        "timestamp": now.isoformat()
    })

# Audio processing status
@app.route('/api/audio-status/<int:audio_id>', methods=['GET'])
def get_audio_status(audio_id):
    status = processing_audios.get(audio_id)
    if status:
        return jsonify(status)
    else:
        return jsonify({"error": "Audio no encontrado o procesamiento completado"}), 404

# Scenario management endpoints
@app.route('/api/escenarios/<int:escenario_id>/iniciar', methods=['POST'])
def start_escenario(escenario_id):
    escenario = Escenario.query.get_or_404(escenario_id)
    if escenario.estado != 'programado':
        return jsonify({'success': False, 'error': 'El escenario ya se ha iniciado o finalizado'}), 400
    
    escenario.estado = 'activo'
    db.session.commit()
    
    now = datetime.datetime.now()
    if escenario.end_time > now:
        scheduler.add_job(
            func=finalizar_escenario_job,
            trigger='date',
            run_date=escenario.end_time,
            args=[escenario.id]
        )
    
    return jsonify({
        'success': True,
        'message': 'Escenario iniciado exitosamente',
        'data': {
            'finalizacion_programada': escenario.end_time.isoformat()
        }
    })

@app.route('/api/escenarios/<int:escenario_id>/finalizar', methods=['POST'])
def stop_escenario(escenario_id):
    escenario = Escenario.query.get_or_404(escenario_id)
    if escenario.estado != 'activo':
        return jsonify({'success': False, 'error': 'El escenario no está en ejecución'}), 400
    
    # Save microphone history
    mic_list = list(escenario.microfonos)
    historial = [m.id for m in mic_list]
    escenario.microfonos_historial = json.dumps(historial)
    escenario.estado = 'culminado'
    
    # Free microphones
    for microfono in mic_list:
        microfono.escenario_id = None
    
    db.session.commit()
    
    return jsonify({
        'success': True,
        'message': 'Escenario finalizado exitosamente'
    })

# Microphone assignment
@app.route('/api/escenarios/<int:escenario_id>/asignar-microfonos', methods=['POST'])
def assign_microphones(escenario_id):
    escenario = Escenario.query.get_or_404(escenario_id)
    if escenario.estado != 'programado':
        return jsonify({'success': False, 'error': 'No se pueden asignar micrófonos a un escenario iniciado'}), 400
    
    data = request.json
    microfono_ids = data.get('microfono_ids', [])
    
    microfonos = Microfono.query.filter(
        Microfono.id.in_(microfono_ids),
        Microfono.escenario_id.is_(None)
    ).all()
    
    if len(microfonos) != len(microfono_ids):
        return jsonify({'success': False, 'error': 'Algunos micrófonos ya están en uso'}), 400
    
    for microfono in microfonos:
        microfono.escenario_id = escenario.id
    
    db.session.commit()
    
    return jsonify({
        'success': True,
        'message': f'{len(microfonos)} micrófonos asignados exitosamente'
    })

# Get scenario details with results
@app.route('/api/escenarios/<int:escenario_id>/detalle', methods=['GET'])
def get_escenario_detail(escenario_id):
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
                'modelo': micro.modelo,
                'ubicacion': micro.ubicacion,
                'resumen': json.loads(resultado.global_result) if resultado.global_result else {},
                'ultimo_audio': resultado.timestamp.isoformat()
            }
            microfonos_data.append(micro_data)
    
    return jsonify({
        'escenario': {
            'id': escenario.id,
            'nombre': escenario.nombre,
            'descripcion': escenario.descripcion,
            'ubicacion': escenario.ubicacion,
            'start_time': escenario.start_time.strftime("%Y-%m-%d %H:%M:%S"),
            'end_time': escenario.end_time.strftime("%Y-%m-%d %H:%M:%S"),
            'estado': escenario.estado,
            'horas_medicion': escenario.horas_medicion,
            'dias_medicion': escenario.dias_medicion
        },
        'microfonos': microfonos_data,
        'total_resultados': len(resultados)
    })

# Analysis endpoints
@app.route('/api/escenarios/<int:escenario_id>/microfono/<int:mic_id>/analisis-excesos', methods=['GET'])
def get_analisis_excesos(escenario_id, mic_id):
    resultados = AudioResultado.query.filter_by(escenario_id=escenario_id, microfono_id=mic_id).all()
    if not resultados:
        return jsonify({"error": "No hay datos para analizar"}), 404
    
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
            
            for punto in detailed_results:
                total_puntos += 1
                lp_max = float(punto.get('Lp_max', 0))
                
                if lp_max > limite_seguro:
                    puntos_excedidos += 1
                    audio_excesos += 1
                    duracion_audio_exceso += 1
                    
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
    
    return jsonify({
        "limite_referencia": limite_seguro,
        "total_puntos_medidos": total_puntos,
        "puntos_que_excedieron": puntos_excedidos,
        "porcentaje_exceso": round(porcentaje_exceso, 2),
        "duracion_total_exceso_segundos": duracion_total_exceso,
        "duracion_total_exceso_minutos": round(duracion_total_exceso / 60, 2),
        "nivel_maximo_registrado": round(max_exceso, 2),
        "cantidad_audios": len(resultados),
        "excesos_por_audio": excesos_por_audio,
        "evaluacion_seguridad": "SEGURO" if puntos_excedidos == 0 else "ATENCIÓN REQUERIDA"
    })

# ============================================================================
# LEGACY ROUTES (keeping for backward compatibility)
# ============================================================================

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/upload_audio', methods=['POST'])
def upload_audio():
    # Keep original upload_audio for Arduino compatibility
    try:
        api_key = request.headers.get('X-API-Key')
        if not api_key or api_key != app.config['API_KEY']:
            return jsonify({"error": "API key requerida o inválida"}), 401

        if 'audio' not in request.files:
            return jsonify({"error": "Campo 'audio' requerido"}), 400

        audio_file = request.files['audio']
        if audio_file.filename == '':
            return jsonify({"error": "Nombre de archivo inválido"}), 400

        filename = secure_filename(audio_file.filename)
        upload_folder = os.path.abspath(app.config['UPLOAD_FOLDER'])
        filepath = os.path.join(upload_folder, filename)

        os.makedirs(upload_folder, exist_ok=True)
        audio_file.save(filepath)

        new_audio = Audio(
            filename=filename,
            title=request.form.get('title', 'Audio desde Arduino'),
            source=request.form.get('source', 'Arduino')
        )
        db.session.add(new_audio)
        db.session.commit()

        socketio.emit('new_audio', {
            'audio_id': new_audio.id,
            'message': 'Nuevo audio cargado'
        }, namespace="/")

        global_result, detailed_results = analizar_audio_file(filepath, audio_id=new_audio.id)

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

        return jsonify({
            "message": "Audio subido y analizado",
            "filename": filename,
            "size": os.path.getsize(filepath),
            "audio_id": new_audio.id,
            "asignacion": indicador_asignacion
        }), 200

    except Exception as e:
        logger.exception("Error durante la subida y análisis del audio")
        return jsonify({"error": str(e)}), 500

# Keep other legacy routes for existing functionality
@app.route('/escenarios', methods=['POST'])
def crear_escenario():
    return create_escenario()

@app.route('/escenarios', methods=['GET'])
def listar_escenarios():
    response = get_escenarios()
    return jsonify(response.get_json()["escenarios"])

@app.route('/micros', methods=['GET'])
def listar_microfonos():
    response = get_microfonos()
    return jsonify(response.get_json()["microfonos"])

@app.route('/micros', methods=['POST'])
def crear_microfono():
    return create_microfono()

@app.route('/dashboard_data')
def dashboard_data():
    response = get_dashboard_data()
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
