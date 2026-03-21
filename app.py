from flask import Flask, render_template, request, redirect, url_for, send_from_directory, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from apscheduler.schedulers.background import BackgroundScheduler
from scipy.io import wavfile
from scipy import signal
from flask_socketio import SocketIO, emit
from flask_cors import CORS
from werkzeug.utils import secure_filename
import soundfile as sf
import os
import logging
import numpy as np
from io import BytesIO
import datetime
import json
import pytz
import math
import traceback
import sys
import base64
import zipfile
import uuid

try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except Exception:
    Document = None
    Inches = None
    Pt = None
    WD_ALIGN_PARAGRAPH = None
    OxmlElement = None
    qn = None

try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
except Exception:
    matplotlib = None
    plt = None

try:
    from openpyxl import Workbook, load_workbook
except Exception:
    Workbook = None
    load_workbook = None

try:
    from docx_generator import build_informe_docx, build_informe_combinado_docx, _image_from_path
    _docx_available = True
except Exception:
    _docx_available = False


# Configurar logger
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
CORS(app)
socketio = SocketIO(app, cors_allowed_origins="*")
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['PHOTOS_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads', 'fotos')
app.config['ALLOWED_IMAGE_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
app.config['API_KEY'] = 'acoustics'
app.config['ALLOWED_IMPORT_AUDIO_EXTENSIONS'] = {'.wav'}
db = SQLAlchemy(app)

# Crear carpetas de subidas si no existen
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])
    logger.debug(f"Carpeta de subidas creada en: {app.config['UPLOAD_FOLDER']}")
os.makedirs(app.config['PHOTOS_FOLDER'], exist_ok=True)

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
LIMA_TZ = pytz.timezone("America/Lima")

def lima_now():
    return datetime.datetime.now(LIMA_TZ)

def to_lima_datetime(dt):
    if dt is None:
        return None
    if dt.tzinfo is None:
        return LIMA_TZ.localize(dt)
    return dt.astimezone(LIMA_TZ)

def lima_iso(dt):
    dt_lima = to_lima_datetime(dt)
    return dt_lima.isoformat() if dt_lima else None

def lima_epoch_ms(dt):
    dt_lima = to_lima_datetime(dt)
    return int(dt_lima.timestamp() * 1000) if dt_lima else 0

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
    fotos = db.Column(db.Text, nullable=True)  # JSON list of photo filenames

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
                "start_time": lima_now().isoformat(),
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
                "start_time": lima_now().isoformat(),
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

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'), 'favicon.svg', mimetype='image/svg+xml')

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


@app.route('/audio/<int:audio_id>/reporte_word', methods=['GET'])
def reporte_word_audio_individual(audio_id):
    deps = _export_dependencies_status()
    if not deps['python_docx']['available']:
        return jsonify({
            'error': 'python-docx no esta disponible en el entorno',
            'hint': 'Ejecuta: python -m pip install -r requirements.txt y reinicia la app',
            'dependencies': deps
        }), 500
    if not deps['matplotlib']['available']:
        return jsonify({
            'error': 'matplotlib no esta disponible en el entorno',
            'hint': 'Ejecuta: python -m pip install -r requirements.txt y reinicia la app',
            'dependencies': deps
        }), 500

    audio = Audio.query.get_or_404(audio_id)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], audio.filename)
    try:
        global_result, detailed_results = analizar_audio_file(filepath, audio_id=audio.id)
    except Exception as e:
        return jsonify({'error': f'No se pudo analizar el audio: {e}'}), 500

    doc = Document()
    _set_doc_base_style(doc)
    doc.add_heading('INFORME TECNICO DE ANALISIS INDIVIDUAL', level=0)
    doc.add_paragraph(f'Audio ID: {audio.id}')
    doc.add_paragraph(f'Titulo: {audio.title}')
    doc.add_paragraph(f'Fuente: {audio.source}')
    doc.add_paragraph(f'Archivo: {audio.filename}')
    doc.add_paragraph(f'Fecha de analisis: {lima_now().strftime("%Y-%m-%d %H:%M:%S")} (America/Lima)')
    doc.add_page_break()

    doc.add_heading('Grafico temporal del audio', level=1)
    chart_buffer = _build_audio_chart_png(detailed_results, audio.id, limite_referencia=SAFE_NOISE_LEVEL)
    doc.add_picture(chart_buffer, width=Inches(6.5))
    doc.add_page_break()

    doc.add_heading('Resumen global editable', level=1)
    _add_key_value_table(doc, 'Resultados globales', global_result)
    doc.add_paragraph('')
    _add_detailed_table(doc, detailed_results)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    safe_name = secure_filename(audio.title) or f'audio_{audio.id}'
    filename = f'informe_word_audio_{audio.id}_{safe_name}.docx'
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

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


def _get_or_create_sensor1_microfono():
    micro = Microfono.query.filter_by(identificador='sensor_1').first()
    if micro:
        return micro
    micro = Microfono(
        identificador='sensor_1',
        modelo='Sensor virtual (importacion ZIP)',
        ubicacion='Importado'
    )
    db.session.add(micro)
    db.session.flush()
    return micro


def _zip_member_scenario_and_file(member_name):
    normalized = str(member_name or '').replace('\\', '/').strip('/')
    if not normalized or normalized.endswith('/'):
        return None, None
    parts = [p for p in normalized.split('/') if p and p != '__MACOSX']
    if not parts:
        return None, None
    if len(parts) == 1:
        scenario_name = 'escenario_importado'
        filename = parts[0]
    else:
        scenario_name = parts[0]
        filename = parts[-1]
    return scenario_name, filename


def _zip_member_path_parts(member_name):
    normalized = str(member_name or '').replace('\\', '/').strip('/')
    if not normalized or normalized.endswith('/'):
        return []
    parts = [p for p in normalized.split('/') if p and p != '__MACOSX']
    return parts


def _naive_lima_now():
    return lima_now().replace(tzinfo=None)


@app.route('/escenarios/importar_zip', methods=['POST'])
def importar_zip_escenarios():
    """Importa un ZIP con estructura escenarios/<archivos.wav> y crea escenarios culminados."""
    if 'zip_file' not in request.files:
        return jsonify({'error': "Campo 'zip_file' requerido"}), 400

    zip_file = request.files['zip_file']
    if not zip_file or not zip_file.filename:
        return jsonify({'error': 'Archivo ZIP invalido'}), 400

    filename = secure_filename(zip_file.filename)
    if not filename.lower().endswith('.zip'):
        return jsonify({'error': 'Solo se admite formato .zip'}), 400

    try:
        zf = zipfile.ZipFile(zip_file.stream)
    except Exception:
        return jsonify({'error': 'No se pudo abrir el ZIP'}), 400

    try:
        grouped = {}
        candidate_members = []
        for info in zf.infolist():
            if info.is_dir():
                continue
            parts = _zip_member_path_parts(info.filename)
            if not parts:
                continue
            inner_file = parts[-1]
            ext = os.path.splitext(inner_file.lower())[1]
            if ext not in app.config['ALLOWED_IMPORT_AUDIO_EXTENSIONS']:
                continue
            candidate_members.append((info, parts))

        if not candidate_members:
            return jsonify({'error': 'No se encontraron audios WAV en el ZIP'}), 400

        # Si todo el ZIP está dentro de una sola carpeta contenedora,
        # se elimina ese prefijo para que cada subcarpeta sea un escenario.
        first_segments = {parts[0] for _, parts in candidate_members if len(parts) >= 2}
        has_nested_paths = any(len(parts) >= 3 for _, parts in candidate_members)
        strip_root_container = len(first_segments) == 1 and has_nested_paths

        for info, parts in candidate_members:
            work_parts = parts[1:] if strip_root_container and len(parts) > 1 else parts
            if len(work_parts) == 1:
                scenario_name = 'escenario_importado'
                inner_file = work_parts[0]
            else:
                scenario_name = work_parts[0]
                inner_file = work_parts[-1]
            grouped.setdefault(scenario_name, []).append((info, inner_file))

        if not grouped:
            return jsonify({'error': 'No se encontraron audios WAV en el ZIP'}), 400

        base_start = _naive_lima_now()
        global_offset_seconds = 0.0
        created_scenarios = []
        total_imported_audios = 0

        sensor1 = _get_or_create_sensor1_microfono()

        for scen_idx, scenario_name in enumerate(sorted(grouped.keys()), start=1):
            members = sorted(grouped[scenario_name], key=lambda m: m[0].filename.lower())

            escenario_start = base_start + datetime.timedelta(seconds=global_offset_seconds)
            provisional_end = escenario_start + datetime.timedelta(seconds=1)
            escenario = Escenario(
                nombre=scenario_name,
                descripcion='Escenario importado masivamente desde ZIP',
                ubicacion='Importado',
                start_time=escenario_start,
                end_time=provisional_end,
                estado='culminado',
                horas_medicion=0,
                dias_medicion=0,
                tipo_ruido='Importado',
                num_fuentes=1,
                num_personas=0,
                proteccion_auditiva='Sí',
                tipo_analisis='Importacion masiva'
            )
            db.session.add(escenario)
            db.session.flush()

            elapsed_seconds = 0.0
            imported_in_scenario = 0

            for audio_idx, member_entry in enumerate(members, start=1):
                member, original_name = member_entry
                try:
                    raw = zf.read(member)
                except Exception:
                    logger.warning(f'No se pudo leer miembro ZIP: {member.filename}')
                    continue

                safe_name = secure_filename(original_name or f'audio_{audio_idx}.wav')
                if not safe_name:
                    safe_name = f'audio_{audio_idx}.wav'

                unique_name = f"imp_{escenario.id}_{audio_idx:04d}_{uuid.uuid4().hex[:8]}_{safe_name}"
                save_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_name)

                with open(save_path, 'wb') as f:
                    f.write(raw)

                audio_timestamp = escenario_start + datetime.timedelta(seconds=elapsed_seconds)
                audio = Audio(
                    filename=unique_name,
                    title=f'{scenario_name} - Audio {audio_idx}',
                    source='Importacion ZIP',
                    timestamp=audio_timestamp
                )
                db.session.add(audio)
                db.session.flush()

                global_result, detailed_results = analizar_audio_file(save_path, audio_id=audio.id)
                resultado = AudioResultado(
                    audio_id=audio.id,
                    microfono_id=sensor1.id,
                    escenario_id=escenario.id,
                    timestamp=audio_timestamp,
                    global_result=json.dumps(global_result),
                    detailed_results=json.dumps(detailed_results)
                )
                db.session.add(resultado)

                duration = float(global_result.get('duration', 0) or 0)
                if duration <= 0:
                    duration = max(float(len(detailed_results or [])), 1.0)

                elapsed_seconds += duration + 0.1
                imported_in_scenario += 1
                total_imported_audios += 1

            if imported_in_scenario == 0:
                db.session.delete(escenario)
                continue

            escenario.end_time = escenario_start + datetime.timedelta(seconds=max(elapsed_seconds - 0.1, 1.0))
            duracion_segundos = (escenario.end_time - escenario.start_time).total_seconds()
            escenario.horas_medicion = round(duracion_segundos / 3600.0, 2)
            escenario.dias_medicion = round(duracion_segundos / (3600.0 * 24), 2)
            escenario.microfonos_historial = json.dumps([sensor1.id])

            created_scenarios.append({
                'escenario_id': escenario.id,
                'nombre': escenario.nombre,
                'audios_importados': imported_in_scenario,
                'inicio': escenario.start_time.strftime('%Y-%m-%d %H:%M:%S'),
                'fin': escenario.end_time.strftime('%Y-%m-%d %H:%M:%S')
            })

            global_offset_seconds += elapsed_seconds + 1.0

        db.session.commit()

        if not created_scenarios:
            return jsonify({'error': 'No se pudieron importar audios validos del ZIP'}), 400

        return jsonify({
            'mensaje': 'Importacion completada',
            'sensor_virtual': {'id': sensor1.id, 'identificador': sensor1.identificador},
            'escenarios_creados': len(created_scenarios),
            'audios_importados': total_imported_audios,
            'detalle': created_scenarios
        }), 200
    except Exception as e:
        db.session.rollback()
        logger.exception('Error importando ZIP de escenarios')
        return jsonify({'error': str(e)}), 500
    finally:
        try:
            zf.close()
        except Exception:
            pass

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
            'start_time': to_lima_datetime(e.start_time).strftime("%Y-%m-%d %H:%M:%S"),
            'end_time': to_lima_datetime(e.end_time).strftime("%Y-%m-%d %H:%M:%S"),
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


def _parse_datetime_flexible(value):
    """Parse datetime strings from UI payloads supporting common formats."""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return to_lima_datetime(value).replace(tzinfo=None)

    text = str(value).strip()
    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M"
    ]
    for fmt in formats:
        try:
            parsed = datetime.datetime.strptime(text, fmt)
            return parsed
        except ValueError:
            continue
    raise ValueError(f"Formato de fecha inválido: {value}")


@app.route('/escenarios/<int:id>', methods=['PUT'])
def actualizar_escenario(id):
    data = request.json
    escenario = Escenario.query.get_or_404(id)
    try:
        if escenario.estado not in ('programado', 'culminado'):
            return jsonify({'error': 'Solo se pueden modificar escenarios programados o culminados'}), 400

        _actualizar_campos_descriptivos(escenario, data)

        if 'start_time' in data:
            escenario.start_time = _parse_datetime_flexible(data['start_time'])
        if 'end_time' in data:
            escenario.end_time = _parse_datetime_flexible(data['end_time'])

        if escenario.start_time and escenario.end_time and escenario.start_time >= escenario.end_time:
            return jsonify({'error': 'La fecha/hora de inicio debe ser menor que la de fin'}), 400

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
    now = lima_now()
    if escenario.end_time > now:
        scheduler.add_job(func=finalizar_escenario_job, trigger='date', run_date=escenario.end_time, args=[escenario.id])
    else:
        finalizar_escenario_job(escenario.id)
    return jsonify({'mensaje': 'Escenario iniciado', 'finalizacion_programada': lima_iso(escenario.end_time)})

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
        'start_time': to_lima_datetime(escenario.start_time).strftime("%Y-%m-%d %H:%M:%S"),
        'end_time': to_lima_datetime(escenario.end_time).strftime("%Y-%m-%d %H:%M:%S"),
        'estado': escenario.estado,
        'horas_medicion': escenario.horas_medicion,
        'dias_medicion': escenario.dias_medicion,
        'tipo_ruido': escenario.tipo_ruido,
        'num_fuentes': escenario.num_fuentes,
        'num_personas': escenario.num_personas,
        'proteccion_auditiva': escenario.proteccion_auditiva,
        'tipo_analisis': escenario.tipo_analisis,
        'microfonos_historial': escenario.microfonos_historial,
        'microfonos': microfonos_data,
        'fotos': json.loads(escenario.fotos) if escenario.fotos else []
    }
    return jsonify(data)


def _allowed_image(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_IMAGE_EXTENSIONS']


MAX_PHOTO_SIZE_BYTES = 10 * 1024 * 1024  # 10 MB per photo


@app.route('/escenarios/<int:escenario_id>/fotos', methods=['POST'])
def subir_foto_escenario(escenario_id):
    """Upload one or more photos for a scenario."""
    escenario = Escenario.query.get_or_404(escenario_id)
    if 'fotos' not in request.files:
        return jsonify({'error': 'No se enviaron archivos'}), 400
    files = request.files.getlist('fotos')
    fotos_actuales = json.loads(escenario.fotos) if escenario.fotos else []
    urls = []
    for foto in files:
        if not foto or not foto.filename:
            continue
        if not _allowed_image(foto.filename):
            continue
        base_filename = secure_filename(foto.filename)
        # secure_filename may return empty string for non-ASCII names
        if not base_filename or base_filename in ('.', '..'):
            # Fallback: use a timestamp-based name with original extension
            ext = foto.filename.rsplit('.', 1)[-1].lower() if '.' in foto.filename else 'jpg'
            if ext not in app.config['ALLOWED_IMAGE_EXTENSIONS']:
                continue
            import time
            base_filename = f"foto_{int(time.time())}.{ext}"
        filename = f"esc{escenario_id}_{base_filename}"
        filepath = os.path.join(app.config['PHOTOS_FOLDER'], filename)
        # Read file data and check size before saving
        foto_data = foto.read()
        if len(foto_data) > MAX_PHOTO_SIZE_BYTES:
            return jsonify({'error': f'El archivo {base_filename} supera el límite de 10 MB'}), 413
        with open(filepath, 'wb') as f:
            f.write(foto_data)
        fotos_actuales.append(filename)
        urls.append(f'/fotos/{filename}')
    escenario.fotos = json.dumps(fotos_actuales)
    db.session.commit()
    return jsonify({'fotos': fotos_actuales, 'urls': urls})


@app.route('/escenarios/<int:escenario_id>/fotos/<filename>', methods=['DELETE'])
def eliminar_foto_escenario(escenario_id, filename):
    """Remove a photo from a scenario."""
    escenario = Escenario.query.get_or_404(escenario_id)
    fotos_actuales = json.loads(escenario.fotos) if escenario.fotos else []
    safe_name = secure_filename(filename)
    if safe_name not in fotos_actuales:
        return jsonify({'error': 'Foto no encontrada'}), 404
    fotos_actuales.remove(safe_name)
    escenario.fotos = json.dumps(fotos_actuales)
    db.session.commit()
    # Delete file from disk
    filepath = os.path.join(app.config['PHOTOS_FOLDER'], safe_name)
    if os.path.exists(filepath):
        os.remove(filepath)
    return jsonify({'mensaje': 'Foto eliminada'})


@app.route('/fotos/<filename>')
def servir_foto(filename):
    """Serve scenario photos."""
    return send_from_directory(app.config['PHOTOS_FOLDER'], filename)


@app.route('/escenarios/culminados', methods=['GET'])
def listar_escenarios_culminados():
    """Return all completed scenarios with their microphones for multi-print selection."""
    escenarios = Escenario.query.filter_by(estado='culminado').order_by(Escenario.end_time.desc()).all()
    resultado = []
    for e in escenarios:
        try:
            mic_ids = json.loads(e.microfonos_historial) if e.microfonos_historial else []
        except Exception:
            mic_ids = []
        mics = []
        for mid in mic_ids:
            m = Microfono.query.get(mid)
            if m:
                mics.append({'id': m.id, 'identificador': m.identificador})
        resultado.append({
            'id': e.id,
            'nombre': e.nombre,
            'ubicacion': e.ubicacion or '',
            'start_time': to_lima_datetime(e.start_time).strftime("%Y-%m-%d %H:%M:%S"),
            'end_time': to_lima_datetime(e.end_time).strftime("%Y-%m-%d %H:%M:%S"),
            'microfonos': mics
        })
    return jsonify(resultado)

@app.route('/micros/<int:microfono_id>/resultado', methods=['GET'])
def detalle_microfono(microfono_id):
    resultado = AudioResultado.query.filter_by(microfono_id=microfono_id).order_by(AudioResultado.timestamp.desc()).first()
    if resultado:
        data = {
            'audio_id': resultado.audio_id,
            'microfono_id': resultado.microfono_id,
            'escenario_id': resultado.escenario_id,
            'timestamp': lima_iso(resultado.timestamp),
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
            "timestamp": lima_iso(res.timestamp),
            "global_result": global_result,
            "detailed_results": detailed_results,
            "filename": filename
        })
    return jsonify(audios)


def _build_audio_chart_png(detailed_results, audio_id, limite_referencia=85.0, lp_eq_global=None):
    """Build a PNG chart for a single audio using detailed points."""
    if plt is None:
        raise RuntimeError('matplotlib no esta disponible en el entorno')

    seconds = []
    lp_eq_values = []
    lp_max_values = []

    for p in detailed_results or []:
        try:
            sec = float(p.get('timestamp', 0))
            seconds.append(sec)
            lp_eq_values.append(float(p.get('Lp_eqT', 0)))
            lp_max_values.append(float(p.get('Lp_max', 0)))
        except Exception:
            continue

    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=160)
    if seconds:
        ax.plot(seconds, lp_eq_values, color='#2563eb', linewidth=2.0, label='Lp_eqT (dB)')
        ax.plot(seconds, lp_max_values, color='#ef4444', linewidth=1.8, alpha=0.9, label='Lp_max (dB)')
    else:
        ax.text(0.5, 0.5, 'Sin datos detallados para este audio', ha='center', va='center', transform=ax.transAxes)

    ax.axhline(y=limite_referencia, color='#dc2626', linestyle='--', linewidth=1.8, label=f'Referencia {limite_referencia} dB')
    if lp_eq_global is not None:
        try:
            lp_eq_global = float(lp_eq_global)
            if np.isfinite(lp_eq_global):
                ax.axhline(y=lp_eq_global, color='#7e22ce', linestyle='-.', linewidth=1.8, label=f'LEQ global audio {lp_eq_global:.2f} dB')
        except Exception:
            pass
    ax.set_title(f'Audio {audio_id} - Analisis temporal por segundo', fontsize=12, pad=10)
    ax.set_xlabel('Tiempo relativo (s)')
    ax.set_ylabel('Nivel (dB)')
    ax.grid(True, linestyle='--', alpha=0.25)
    ax.legend(loc='best', fontsize=8)
    fig.tight_layout()

    buffer = BytesIO()
    fig.savefig(buffer, format='png')
    plt.close(fig)
    buffer.seek(0)
    return buffer


def _parse_bool_param(name, default=True):
    value = request.args.get(name)
    if value is None:
        return default
    return str(value).strip().lower() in ('1', 'true', 'yes', 'si', 'on')


def _export_dependencies_status():
    return {
        'python_docx': {
            'available': Document is not None and Inches is not None,
            'detail': 'ok' if (Document is not None and Inches is not None) else 'python-docx no cargado en runtime'
        },
        'matplotlib': {
            'available': plt is not None,
            'detail': 'ok' if plt is not None else 'matplotlib no cargado en runtime'
        },
        'openpyxl': {
            'available': Workbook is not None and load_workbook is not None,
            'detail': 'ok' if (Workbook is not None and load_workbook is not None) else 'openpyxl no cargado en runtime'
        }
    }


@app.route('/diagnostico/exportes', methods=['GET'])
def diagnostico_exportes():
    deps = _export_dependencies_status()
    word_ok = deps['python_docx']['available'] and deps['matplotlib']['available']
    return jsonify({
        'word_export_available': word_ok,
        'dependencies': deps,
        'recommendation': 'Ejecuta: python -m pip install -r requirements.txt y reinicia la app' if not word_ok else 'Entorno listo para exportar Word'
    })


def _label_global_metric(key):
    labels = {
        'Lp_eqT': 'Nivel equivalente (Lp_eqT, dB)',
        'Lp_max': 'Nivel maximo (Lp_max, dB)',
        'ET': 'Energia total (ET)',
        'LE': 'Nivel de energia (LE, dB)',
        'J_energy': 'Energia J',
        'LJ': 'Nivel de energia LJ (dB)',
        'IT': 'Intensidad IT',
        'LI': 'Nivel de intensidad LI (dB)',
        'duration': 'Duracion (s)',
        'sample_rate': 'Frecuencia de muestreo (Hz)',
        'alert': 'Alerta'
    }
    return labels.get(key, str(key))


def _set_doc_base_style(doc):
    if Pt is None:
        return
    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(10.5)


def _set_table_header_cell(cell, text, fill='1E293B', color='FFFFFF'):
    cell.text = str(text)
    if OxmlElement is None or qn is None:
        return
    try:
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill)
        tc_pr.append(shd)
        for p in cell.paragraphs:
            for r in p.runs:
                if r.font is not None:
                    r.font.bold = True
                    r.font.color.rgb = None
    except Exception:
        return


def _add_doc_cover(doc, escenario, microfono, total_audios):
    title = doc.add_heading('INFORME TECNICO DE MONITOREO ACUSTICO', level=0)
    if WD_ALIGN_PARAGRAPH is not None:
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph('Resultados por audio y analisis detallado por escenario culminado')
    if WD_ALIGN_PARAGRAPH is not None:
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph('')
    meta = doc.add_table(rows=0, cols=2)
    meta.style = 'Table Grid'
    rows = [
        ('Escenario', f'{escenario.nombre} (ID {escenario.id})'),
        ('Microfono', f'{microfono.identificador} (ID {microfono.id})'),
        ('Estado', 'Culminado'),
        ('Total de audios analizados', str(total_audios)),
        ('Generado en', f"{lima_now().strftime('%Y-%m-%d %H:%M:%S')} (America/Lima)"),
        ('Limite de referencia', f'{SAFE_NOISE_LEVEL} dB')
    ]
    for k, v in rows:
        r = meta.add_row().cells
        r[0].text = k
        r[1].text = v


def _add_key_value_table(doc, title, values):
    doc.add_heading(title, level=2)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    _set_table_header_cell(hdr[0], 'Metrica')
    _set_table_header_cell(hdr[1], 'Valor')
    for key, value in values.items():
        row = table.add_row().cells
        row[0].text = _label_global_metric(key)
        row[1].text = str(value)

def _add_detailed_table(doc, detailed_results, chunk_size=45):
    headers = ['Segundo', 'Lp_eqT', 'Lp_max', 'ET', 'LE', 'J_energy', 'LJ', 'IT', 'LI', 'Alerta']
    rows = detailed_results or []
    if not rows:
        doc.add_paragraph('No hay datos detallados por segundo para este audio.')
        return

    parts = [rows[i:i + chunk_size] for i in range(0, len(rows), chunk_size)]
    for part_idx, part in enumerate(parts, start=1):
        if len(parts) > 1:
            doc.add_heading(f'Datos detallados por segundo (parte {part_idx}/{len(parts)})', level=2)
        else:
            doc.add_heading('Datos detallados por segundo', level=2)

        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        for idx, header in enumerate(headers):
            _set_table_header_cell(table.rows[0].cells[idx], header)

        for p in part:
            row = table.add_row().cells
            row[0].text = str(p.get('timestamp', ''))
            row[1].text = str(p.get('Lp_eqT', ''))
            row[2].text = str(p.get('Lp_max', ''))
            row[3].text = str(p.get('ET', ''))
            row[4].text = str(p.get('LE', ''))
            row[5].text = str(p.get('J_energy', ''))
            row[6].text = str(p.get('LJ', ''))
            row[7].text = str(p.get('IT', ''))
            row[8].text = str(p.get('LI', ''))
            row[9].text = str(p.get('alert', ''))

        if part_idx < len(parts):
            doc.add_page_break()


def _compute_general_analysis_from_results(resultados):
    total_duration = 0.0
    total_ET = 0.0
    weighted_sum_pressure_sq = 0.0
    for res in resultados:
        try:
            global_result = json.loads(res.global_result) if res.global_result else {}
        except Exception:
            global_result = {}

        duration = float(global_result.get('duration', 0) or 0)
        lp_eqt = float(global_result.get('Lp_eqT', 0) or 0)
        et = float(global_result.get('ET', 0) or 0)

        pressure_sq = (2.0e-5) ** 2 * (10 ** (lp_eqt / 10)) if lp_eqt else 0
        total_duration += duration
        total_ET += et
        weighted_sum_pressure_sq += pressure_sq * duration

    mean_pressure_sq = weighted_sum_pressure_sq / total_duration if total_duration > 0 else 0
    lp_eqt_global = 10 * np.log10(mean_pressure_sq / (2.0e-5) ** 2) if mean_pressure_sq > 0 else 0
    le_global = 10 * np.log10(total_ET / 4.0e-10) if total_ET > 0 else 0
    duracion_horas = total_duration / 3600.0
    l_ex_8h = (lp_eqt_global + 10 * np.log10(duracion_horas / 8.0)) if duracion_horas >= (1.0 / 60.0) else None

    return {
        'duration': round(total_duration, 2),
        'ET': round(total_ET, 6),
        'Lp_eqT': round(lp_eqt_global, 2),
        'LE': round(le_global, 2),
        'L_EX_8h': round(l_ex_8h, 2) if l_ex_8h is not None else 'No aplica',
        'cantidad_audios': len(resultados)
    }


def _build_principal_chart_png(resultados, laeq_global=None, limite_referencia=85.0, escenario_id=None, microfono_id=None):
    if plt is None:
        raise RuntimeError('matplotlib no esta disponible en el entorno')

    # Prioriza la misma serie temporal usada por la vista web del informe final.
    x_seconds = []
    y_lpmax = []

    if escenario_id is not None and microfono_id is not None:
        try:
            client = app.test_client()
            resp = client.get(
                f'/escenarios/{int(escenario_id)}/microfono/{int(microfono_id)}/serie_temporal?max_points=1800'
            )
            if resp.status_code == 200:
                payload = resp.get_json(silent=True) or {}
                points = payload.get('points') or []
                # Mantener TODOS los puntos incluyendo None para discontinuidades entre audios
                base_x = None
                for p in points:
                    x = p.get('x')
                    y = p.get('y')
                    if x is not None:
                        if base_x is None:
                            base_x = x
                        x_seconds.append((x - base_x) / 1000.0)
                        y_lpmax.append(y)  # y puede ser None, matplotlib maneja eso
        except Exception:
            x_seconds = []
            y_lpmax = []

    # Fallback: reconstruye con resultados detallados si no se pudo leer la serie del informe web.
    if not x_seconds:
        offset = 0.0
        for res in resultados:
            try:
                global_result = json.loads(res.global_result) if res.global_result else {}
            except Exception:
                global_result = {}
            try:
                detailed_results = json.loads(res.detailed_results) if res.detailed_results else []
            except Exception:
                detailed_results = []

            local_max_t = 0.0
            for point in detailed_results:
                try:
                    t = float(point.get('timestamp', 0) or 0)
                    lp_max = float(point.get('Lp_max', 0) or 0)
                except Exception:
                    continue
                x_seconds.append(offset + t)
                y_lpmax.append(lp_max)
                if t > local_max_t:
                    local_max_t = t

            duration = float(global_result.get('duration', 0) or 0)
            if duration <= 0:
                duration = local_max_t if local_max_t > 0 else 1.0
            offset += duration

    fig, ax = plt.subplots(figsize=(14, 7), dpi=220)
    if x_seconds:
        ax.plot(x_seconds, y_lpmax, color='#1d4ed8', linewidth=2.2, label='Lp,max (escenario completo)', zorder=3)
    else:
        ax.text(0.5, 0.5, 'Sin datos para grafico principal del escenario', ha='center', va='center', transform=ax.transAxes)

    ax.axhline(y=limite_referencia, color='#dc2626', linestyle='--', linewidth=2.2, label=f'LMP {limite_referencia} dB', zorder=2)
    if laeq_global is not None:
        try:
            laeq = float(laeq_global)
            if np.isfinite(laeq):
                ax.axhline(y=laeq, color='#7e22ce', linestyle='-.', linewidth=2.2, label=f'LAeq general {laeq:.2f} dB', zorder=2)
        except Exception:
            pass

    ax.set_title('Informe final principal: Lp,max vs tiempo del escenario', fontsize=16, fontweight='bold', pad=15)
    ax.set_xlabel('Tiempo acumulado del escenario (s)', fontsize=12, fontweight='bold')
    ax.set_ylabel('Nivel de Presión Sonora (dB)', fontsize=12, fontweight='bold')
    ax.grid(True, linestyle='-', alpha=0.3, linewidth=0.7)
    ax.legend(loc='best', fontsize=11, framealpha=0.95)
    fig.tight_layout()

    buffer = BytesIO()
    fig.savefig(buffer, format='png', dpi=220, bbox_inches='tight', pad_inches=0.15)
    plt.close(fig)
    buffer.seek(0)
    return buffer


def _add_principal_technical_table(doc, resultados, limite_referencia=85.0):
    headers = ['Audio', 'Fecha/Hora', 'LAeq (dB)', 'Lp,max (dB)', 'Duracion (s)', 'Sample Rate (Hz)', 'ET', 'LE (dB)', 'LI (dB)', 'LJ (dB)', 'Estado']
    doc.add_heading('Tabla tecnica consolidada del informe principal', level=2)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    for i, h in enumerate(headers):
        _set_table_header_cell(table.rows[0].cells[i], h)

    for res in resultados:
        try:
            g = json.loads(res.global_result) if res.global_result else {}
        except Exception:
            g = {}

        lp_max = float(g.get('Lp_max', 0) or 0)
        status = f'Excede {limite_referencia} dB' if lp_max > limite_referencia else 'Dentro del limite'
        ts = to_lima_datetime(res.timestamp).strftime('%Y-%m-%d %H:%M:%S') if res.timestamp else 'N/A'

        row = table.add_row().cells
        row[0].text = str(res.audio_id)
        row[1].text = ts
        row[2].text = str(g.get('Lp_eqT', 'N/A'))
        row[3].text = str(g.get('Lp_max', 'N/A'))
        row[4].text = str(g.get('duration', 'N/A'))
        row[5].text = str(g.get('sample_rate', 'N/A'))
        et_val = g.get('ET', 'N/A')
        if isinstance(et_val, (int, float)):
            et_val = f'{float(et_val):.2e}'
        row[6].text = str(et_val)
        row[7].text = str(g.get('LE', 'N/A'))
        row[8].text = str(g.get('LI', 'N/A'))
        row[9].text = str(g.get('LJ', 'N/A'))
        row[10].text = status


def _build_excess_summary(resultados, limite_referencia=85.0):
    total_points = 0
    exceeded_points = 0
    max_level = 0.0
    per_audio = []

    for res in resultados:
        try:
            g = json.loads(res.global_result) if res.global_result else {}
        except Exception:
            g = {}
        try:
            d = json.loads(res.detailed_results) if res.detailed_results else []
        except Exception:
            d = []

        audio_exceeded = 0
        for p in d:
            try:
                val = float(p.get('Lp_max', 0) or 0)
            except Exception:
                continue
            total_points += 1
            if val > limite_referencia:
                exceeded_points += 1
                audio_exceeded += 1
            if val > max_level:
                max_level = val

        per_audio.append({
            'audio_id': res.audio_id,
            'timestamp': to_lima_datetime(res.timestamp).strftime('%Y-%m-%d %H:%M:%S') if res.timestamp else 'N/A',
            'laeq': g.get('Lp_eqT', 'N/A'),
            'lpmax': g.get('Lp_max', 'N/A'),
            'puntos_excedidos': audio_exceeded,
            'estado': 'Excede' if audio_exceeded > 0 else 'OK'
        })

    percent = (exceeded_points / total_points * 100.0) if total_points > 0 else 0.0
    return {
        'limite_referencia_db': limite_referencia,
        'total_puntos_medidos': total_points,
        'puntos_excedidos': exceeded_points,
        'porcentaje_excedido': round(percent, 2),
        'nivel_maximo_registrado': round(max_level, 2),
        'evaluacion': 'ATENCION REQUERIDA' if exceeded_points > 0 else 'DENTRO DEL LIMITE'
    }, per_audio


def _add_excess_detail_table(doc, per_audio_rows):
    doc.add_heading('Detalle de excedencias por audio', level=2)
    headers = ['Audio', 'Fecha/Hora', 'LAeq (dB)', 'Lp,max (dB)', 'Puntos excedidos', 'Estado']
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    for i, h in enumerate(headers):
        _set_table_header_cell(table.rows[0].cells[i], h)

    for item in per_audio_rows:
        row = table.add_row().cells
        row[0].text = str(item.get('audio_id', 'N/A'))
        row[1].text = str(item.get('timestamp', 'N/A'))
        row[2].text = str(item.get('laeq', 'N/A'))
        row[3].text = str(item.get('lpmax', 'N/A'))
        row[4].text = str(item.get('puntos_excedidos', 0))
        row[5].text = str(item.get('estado', 'N/A'))


def _build_scenario_summary(escenario, resultados):
    total_duration = 0.0
    total_ET = 0.0
    weighted_pressure_sq = 0.0
    total_points = 0
    points_exceeded = 0
    peak_max = 0.0

    for res in resultados:
        try:
            g = json.loads(res.global_result) if res.global_result else {}
        except Exception:
            g = {}
        try:
            d = json.loads(res.detailed_results) if res.detailed_results else []
        except Exception:
            d = []

        duration = float(g.get('duration', 0) or 0)
        lp_eq = float(g.get('Lp_eqT', 0) or 0)
        et = float(g.get('ET', 0) or 0)
        pressure_sq = (2.0e-5) ** 2 * (10 ** (lp_eq / 10)) if lp_eq else 0.0

        total_duration += duration
        total_ET += et
        weighted_pressure_sq += pressure_sq * duration

        for p in d:
            try:
                lp_max = float(p.get('Lp_max', 0) or 0)
            except Exception:
                continue
            total_points += 1
            if lp_max > SAFE_NOISE_LEVEL:
                points_exceeded += 1
            if lp_max > peak_max:
                peak_max = lp_max

    mean_pressure_sq = weighted_pressure_sq / total_duration if total_duration > 0 else 0
    lp_eq_global = 10 * np.log10(mean_pressure_sq / (2.0e-5) ** 2) if mean_pressure_sq > 0 else 0
    le_global = 10 * np.log10(total_ET / 4.0e-10) if total_ET > 0 else 0
    dur_hours = total_duration / 3600.0 if total_duration > 0 else 0
    l_ex_8h = (lp_eq_global + 10 * np.log10(dur_hours / 8.0)) if dur_hours >= (1.0 / 60.0) else None
    exceed_pct = (points_exceeded / total_points * 100.0) if total_points > 0 else 0.0

    duration_hours = ((total_duration / 3600.0) if total_duration > 0 else 0.0)
    photo_count = 0
    if escenario.fotos:
        try:
            parsed_photos = json.loads(escenario.fotos)
            if isinstance(parsed_photos, list):
                photo_count = len(parsed_photos)
        except Exception:
            photo_count = 0

    historial_txt = 'No registrado'
    if escenario.microfonos_historial:
        try:
            hist = json.loads(escenario.microfonos_historial)
            if isinstance(hist, list) and hist:
                etiquetas = []
                for mic_id in hist:
                    try:
                        mic = Microfono.query.get(int(mic_id))
                    except Exception:
                        mic = None
                    if mic and mic.identificador:
                        etiquetas.append(mic.identificador)
                    else:
                        etiquetas.append(f'ID {mic_id}')
                historial_txt = ', '.join(etiquetas)
        except Exception:
            historial_txt = str(escenario.microfonos_historial)

    def _opt_text(v):
        if v is None:
            return 'No registrado'
        s = str(v).strip()
        return s if s else 'No registrado'

    scenario_rows = {
        'Escenario ID': escenario.id,
        'Nombre': escenario.nombre,
        'Descripcion': _opt_text(escenario.descripcion),
        'Ubicacion': _opt_text(escenario.ubicacion),
        'Estado': escenario.estado,
        'Inicio': to_lima_datetime(escenario.start_time).strftime('%Y-%m-%d %H:%M:%S') if escenario.start_time else 'N/A',
        'Fin': to_lima_datetime(escenario.end_time).strftime('%Y-%m-%d %H:%M:%S') if escenario.end_time else 'N/A',
        'Duracion total medida (h)': round(duration_hours, 2),
        'Horas de medicion (declaradas)': _opt_text(escenario.horas_medicion),
        'Dias de medicion (declarados)': _opt_text(escenario.dias_medicion),
        'Tipo de ruido': _opt_text(escenario.tipo_ruido),
        'Tipo de analisis': _opt_text(escenario.tipo_analisis),
        'Proteccion auditiva': _opt_text(escenario.proteccion_auditiva),
        'Fuentes declaradas': escenario.num_fuentes if escenario.num_fuentes is not None else 'No registrado',
        'Personas expuestas': escenario.num_personas if escenario.num_personas is not None else 'No registrado',
        'Microfonos (historial)': historial_txt,
        'Fotos adjuntas': photo_count
    }

    kpis = {
        'Audios procesados': len(resultados),
        'Duracion total (s)': round(total_duration, 2),
        'LAeq global (dB)': round(lp_eq_global, 2),
        'LE global (dB)': round(le_global, 2),
        'L_EX,8h (dB)': round(l_ex_8h, 2) if l_ex_8h is not None else 'No aplica',
        'Nivel maximo registrado (dB)': round(peak_max, 2),
        'Puntos sobre 85 dB': points_exceeded,
        'Porcentaje sobre 85 dB': f'{round(exceed_pct, 2)} %'
    }

    return scenario_rows, kpis


@app.route('/escenarios/<int:escenario_id>/microfono/<int:microfono_id>/reporte_word', methods=['GET'])
def reporte_word_por_audio(escenario_id, microfono_id):
    """Generate an editable DOCX report for all audios of a scenario/microphone."""
    deps = _export_dependencies_status()
    if not deps['python_docx']['available']:
        return jsonify({
            'error': 'python-docx no esta disponible en el entorno',
            'hint': 'Ejecuta: python -m pip install -r requirements.txt y reinicia la app',
            'dependencies': deps
        }), 500
    if not deps['matplotlib']['available']:
        return jsonify({
            'error': 'matplotlib no esta disponible en el entorno',
            'hint': 'Ejecuta: python -m pip install -r requirements.txt y reinicia la app',
            'dependencies': deps
        }), 500

    escenario = Escenario.query.get_or_404(escenario_id)
    microfono = Microfono.query.get_or_404(microfono_id)

    include_chart = _parse_bool_param('include_chart', default=True)
    include_global = _parse_bool_param('include_global', default=True)
    include_detailed = _parse_bool_param('include_detailed', default=True)
    include_metodologia = _parse_bool_param('include_metodologia', default=False)
    include_fotos = _parse_bool_param('include_fotos', default=False)
    include_glosario = _parse_bool_param('include_glosario', default=False)
    include_referencias = _parse_bool_param('include_referencias', default=False)
    # Regla de negocio: el informe final principal SIEMPRE se incluye.
    include_final_summary = True

    audio_ids_raw = (request.args.get('audio_ids') or '').strip()
    selected_audio_ids = set()
    if audio_ids_raw:
        for token in audio_ids_raw.split(','):
            token = token.strip()
            if not token:
                continue
            try:
                selected_audio_ids.add(int(token))
            except Exception:
                continue

    resultados_all = AudioResultado.query.filter_by(
        escenario_id=escenario_id,
        microfono_id=microfono_id
    ).order_by(AudioResultado.timestamp.asc()).all()

    if not resultados_all:
        return jsonify({'error': 'No hay audios analizados para este escenario y microfono'}), 404

    # Los audios seleccionados aplican SOLO a secciones individuales.
    resultados = resultados_all
    if selected_audio_ids:
        resultados = [r for r in resultados_all if int(r.audio_id) in selected_audio_ids]

    doc = Document()
    _set_doc_base_style(doc)
    if include_final_summary:
        _add_doc_cover(doc, escenario, microfono, len(resultados_all))
        if selected_audio_ids:
            doc.add_paragraph(f'Modo de seleccion individual: manual ({len(resultados)} audio(s) incluidos)')
        else:
            doc.add_paragraph('Modo de seleccion individual: todos los audios del escenario para este microfono')
        doc.add_paragraph('El informe final principal del escenario se incluye siempre con la totalidad de audios.')

        scenario_rows, kpis = _build_scenario_summary(escenario, resultados_all)
        _add_key_value_table(doc, 'Cuadro 1. Datos del escenario', scenario_rows)
        doc.add_paragraph('')
        _add_key_value_table(doc, 'Cuadro 2. Indicadores globales del escenario', kpis)
        doc.add_page_break()

    doc.add_heading('Indice de contenido', level=1)
    if include_final_summary:
        doc.add_paragraph('0. Informe final principal del escenario (resumen + grafica + tabla tecnica)')
    for idx, res in enumerate(resultados, start=1):
        ts = to_lima_datetime(res.timestamp).strftime('%Y-%m-%d %H:%M:%S') if res.timestamp else 'N/A'
        doc.add_paragraph(f'{idx}. Audio ID {res.audio_id} - {ts}')
    doc.add_page_break()

    if include_final_summary:
        # El bloque principal SIEMPRE usa todos los audios del escenario (igual que el informe web/PDF).
        general = _compute_general_analysis_from_results(resultados_all)
        excess_summary, excess_rows = _build_excess_summary(resultados_all, limite_referencia=SAFE_NOISE_LEVEL)
        doc.add_heading('Informe final principal del escenario', level=1)
        doc.add_paragraph('Esta seccion replica el bloque principal del informe PDF/web: grafica global temporal, resumen tecnico, excedencias y tablas consolidadas del escenario completo.')

        if include_chart:
            chart_principal = _build_principal_chart_png(
                resultados_all,
                laeq_global=general.get('Lp_eqT'),
                limite_referencia=SAFE_NOISE_LEVEL,
                escenario_id=escenario_id,
                microfono_id=microfono_id
            )
            doc.add_paragraph('Grafica principal del escenario (Lp,max temporal con lineas de LAeq general y LMP):')
            doc.add_picture(chart_principal, width=Inches(7.5))
            doc.add_paragraph('')

        if include_global:
            _add_key_value_table(doc, 'Resumen global del informe principal', {
                'Duracion total (s)': general.get('duration', 'N/A'),
                'ET total': general.get('ET', 'N/A'),
                'LAeq general (dB)': general.get('Lp_eqT', 'N/A'),
                'LE global (dB)': general.get('LE', 'N/A'),
                'L_EX,8h (dB)': general.get('L_EX_8h', 'N/A'),
                'Cantidad de audios': general.get('cantidad_audios', len(resultados_all))
            })
            doc.add_paragraph('')
            _add_key_value_table(doc, 'Resumen de excedencias del informe principal', {
                'Limite de referencia (dB)': excess_summary.get('limite_referencia_db'),
                'Total puntos medidos': excess_summary.get('total_puntos_medidos'),
                'Puntos excedidos': excess_summary.get('puntos_excedidos'),
                'Porcentaje excedido (%)': excess_summary.get('porcentaje_excedido'),
                'Nivel maximo registrado (dB)': excess_summary.get('nivel_maximo_registrado'),
                'Evaluacion': excess_summary.get('evaluacion')
            })
            doc.add_paragraph('')
            conclusion = (
                'Conclusion principal: Se identifican excedencias del limite de referencia en el escenario.'
                if excess_summary.get('puntos_excedidos', 0) > 0
                else 'Conclusion principal: No se identifican excedencias del limite de referencia en el escenario.'
            )
            doc.add_paragraph(conclusion)

        if include_detailed:
            _add_principal_technical_table(doc, resultados_all, limite_referencia=SAFE_NOISE_LEVEL)
            doc.add_paragraph('')
            _add_excess_detail_table(doc, excess_rows)

        doc.add_page_break()

    for idx, res in enumerate(resultados, start=1):
        try:
            global_result = json.loads(res.global_result) if res.global_result else {}
        except Exception:
            global_result = {}
        try:
            detailed_results = json.loads(res.detailed_results) if res.detailed_results else []
        except Exception:
            detailed_results = []

        doc.add_heading(f'Audio {idx} de {len(resultados)} - ID {res.audio_id}', level=1)
        doc.add_paragraph(f"Timestamp de analisis: {to_lima_datetime(res.timestamp).strftime('%Y-%m-%d %H:%M:%S')}")
        doc.add_paragraph(f'Escenario: {escenario.nombre} (ID {escenario.id})')
        doc.add_paragraph(f'Microfono: {microfono.identificador} (ID {microfono.id})')

        if include_chart:
            lp_eq_global_audio = global_result.get('Lp_eqT', None) if isinstance(global_result, dict) else None
            chart_buffer = _build_audio_chart_png(
                detailed_results,
                res.audio_id,
                limite_referencia=SAFE_NOISE_LEVEL,
                lp_eq_global=lp_eq_global_audio
            )
            doc.add_paragraph('Grafico del analisis temporal (imagen incrustada):')
            doc.add_picture(chart_buffer, width=Inches(6.5))
            if include_global or include_detailed:
                doc.add_page_break()

        if include_global or include_detailed:
            doc.add_heading(f'Datos editables - Audio ID {res.audio_id}', level=1)
            if include_global:
                _add_key_value_table(doc, 'Resumen global', global_result)
                doc.add_paragraph('')
            if include_detailed:
                _add_detailed_table(doc, detailed_results)

        if idx < len(resultados):
            doc.add_page_break()

    if include_metodologia:
        doc.add_page_break()
        doc.add_heading('Metodologia aplicada', level=1)
        doc.add_paragraph('El procesamiento calcula niveles globales y detallados por segundo a partir de la señal de audio normalizada, incluyendo indicadores de energia e intensidad sonora, con umbral de referencia ocupacional de 85 dB para alertas de excedencia.')

    if include_glosario:
        doc.add_page_break()
        doc.add_heading('Glosario tecnico', level=1)
        glossary = {
            'LAeq': 'Nivel continuo equivalente de presion sonora durante el periodo analizado.',
            'Lp,max': 'Nivel maximo instantaneo registrado en el periodo.',
            'LE': 'Nivel de exposicion sonora acumulada.',
            'L_EX,8h': 'Nivel de exposicion normalizado a 8 horas laborales.',
            'ET': 'Energia acustica total integrada en el tiempo.'
        }
        _add_key_value_table(doc, 'Terminos y definiciones', glossary)

    if include_referencias:
        doc.add_page_break()
        doc.add_heading('Referencias tecnicas', level=1)
        refs = {
            'ISO/TR 25417:2007': 'Definiciones de cantidades y terminos acusticos.',
            'NTP ISO 9612:2010': 'Determinacion de la exposicion al ruido en el trabajo.',
            'D.S. N° 005-2012-TR (Peru)': 'Reglamento de Seguridad y Salud en el Trabajo.'
        }
        _add_key_value_table(doc, 'Normativa aplicada', refs)

    if include_fotos:
        fotos = json.loads(escenario.fotos) if escenario.fotos else []
        fotos_validas = []
        for foto_name in fotos:
            fp = os.path.join(app.config['PHOTOS_FOLDER'], foto_name)
            if os.path.isfile(fp):
                fotos_validas.append(fp)
        if fotos_validas:
            doc.add_page_break()
            doc.add_heading('Registro fotografico del escenario', level=1)
            for i, fp in enumerate(fotos_validas, start=1):
                doc.add_paragraph(f'Foto {i}')
                try:
                    doc.add_picture(fp, width=Inches(5.8))
                except Exception:
                    doc.add_paragraph('No se pudo incrustar una imagen del escenario.')

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    safe_esc = secure_filename(escenario.nombre) or f'escenario_{escenario.id}'
    filename = f'informe_word_{safe_esc}_mic_{microfono.id}.docx'
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/escenarios/<int:escenario_id>/microfono/<int:microfono_id>/serie_temporal', methods=['GET'])
def obtener_serie_temporal(escenario_id, microfono_id):
    resultados = AudioResultado.query.filter_by(
        escenario_id=escenario_id,
        microfono_id=microfono_id
    ).order_by(AudioResultado.timestamp.asc()).all()

    if not resultados:
        return jsonify({
            "points": [],
            "meta": {
                "is_aggregated": False,
                "total_raw_points": 0,
                "returned_points": 0,
                "bucket_ms": 1000,
                "resolution_seconds": 1
            }
        })

    raw_points = []
    for res in resultados:
        try:
            detailed_results = json.loads(res.detailed_results) if res.detailed_results else []
        except Exception:
            detailed_results = []

        base_ms = lima_epoch_ms(res.timestamp)
        for point in detailed_results:
            try:
                rel_seconds = float(point.get('timestamp', 0))
                lp_max = float(point.get('Lp_max', 0))
            except Exception:
                continue
            x_ms = base_ms + int(rel_seconds * 1000)
            raw_points.append((x_ms, lp_max, res.audio_id))

    if not raw_points:
        return jsonify({
            "points": [],
            "meta": {
                "is_aggregated": False,
                "total_raw_points": 0,
                "returned_points": 0,
                "bucket_ms": 1000,
                "resolution_seconds": 1
            }
        })

    raw_points.sort(key=lambda p: p[0])
    global_min = raw_points[0][0]
    global_max = raw_points[-1][0]

    start_ms = request.args.get('start_ms', type=int)
    end_ms = request.args.get('end_ms', type=int)
    max_points = request.args.get('max_points', default=1800, type=int)

    if start_ms is None:
        start_ms = global_min
    if end_ms is None:
        end_ms = global_max
    if end_ms <= start_ms:
        end_ms = start_ms + 1000

    max_points = max(300, min(max_points, 8000))

    filtered = [p for p in raw_points if start_ms <= p[0] <= end_ms]
    total_raw = len(filtered)
    range_fallback = False

    if total_raw == 0:
        filtered = raw_points
        total_raw = len(filtered)
        start_ms = global_min
        end_ms = global_max
        range_fallback = True

    range_ms = max(1, end_ms - start_ms)
    is_aggregated = total_raw > max_points
    bucket_ms = 1000

    if not is_aggregated:
        points = []
        prev_audio = None
        prev_x = None
        for (x, y, aid) in filtered:
            if prev_audio is not None and aid != prev_audio:
                points.append({
                    "x": prev_x + 1 if prev_x is not None else x,
                    "y": None,
                    "audio_id": None,
                    "count": 0,
                    "y_max": None
                })
            points.append({
                "x": x,
                "y": round(y, 4),
                "audio_id": aid,
                "count": 1,
                "y_max": round(y, 4)
            })
            prev_audio = aid
            prev_x = x
    else:
        bucket_ms = max(1, int(math.ceil(range_ms / max_points)))
        buckets = {}
        for x, y, aid in filtered:
            bidx = (x - start_ms) // bucket_ms
            bstart = start_ms + (bidx * bucket_ms)
            b = buckets.get(bstart)
            if not b:
                b = {
                    "sum": 0.0,
                    "count": 0,
                    "y_max": float('-inf'),
                    "audio_counts": {}
                }
                buckets[bstart] = b
            b["sum"] += y
            b["count"] += 1
            if y > b["y_max"]:
                b["y_max"] = y
            b["audio_counts"][aid] = b["audio_counts"].get(aid, 0) + 1

        points = []
        prev_center = None
        prev_audio = None
        for bstart in sorted(buckets.keys()):
            b = buckets[bstart]
            dominant_audio = max(b["audio_counts"], key=b["audio_counts"].get)
            avg_y = b["sum"] / b["count"] if b["count"] else 0.0
            center_x = int(bstart + (bucket_ms // 2))

            if prev_center is not None:
                has_gap = (center_x - prev_center) > int(bucket_ms * 1.5)
                audio_changed = dominant_audio != prev_audio
                if has_gap or audio_changed:
                    points.append({
                        "x": prev_center + 1,
                        "y": None,
                        "audio_id": None,
                        "count": 0,
                        "y_max": None
                    })

            points.append({
                "x": center_x,
                "y": round(avg_y, 4),
                "audio_id": dominant_audio,
                "count": b["count"],
                "y_max": round(b["y_max"], 4)
            })
            prev_center = center_x
            prev_audio = dominant_audio

    return jsonify({
        "points": points,
        "meta": {
            "is_aggregated": is_aggregated,
            "total_raw_points": total_raw,
            "returned_points": len(points),
            "bucket_ms": bucket_ms,
            "resolution_seconds": round(bucket_ms / 1000.0, 3),
            "start_ms": start_ms,
            "end_ms": end_ms,
            "global_min": global_min,
            "global_max": global_max,
            "range_fallback": range_fallback
        }
    })

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
            escenario_start=lima_iso(escenario.start_time),
            escenario_end=lima_iso(escenario.end_time),
            limite_referencia=85.0  # Línea de referencia en 85 dB
        )
    else:
        # Si está activo, usar el template original
        return render_template(
            'monitoring.html',
            escenario_id=escenario_id,
            microfono_id=microfono_id,
            escenario_start=lima_iso(escenario.start_time),
            escenario_end=lima_iso(escenario.end_time)
        )


@app.route('/monitoring_por_audio/<int:escenario_id>/microfono/<int:microfono_id>')
def monitoring_por_audio(escenario_id, microfono_id):
    """Vista alternativa: muestra el analisis de cada audio por separado."""
    escenario = Escenario.query.get_or_404(escenario_id)
    return render_template(
        'monitoring_by_audio.html',
        escenario_id=escenario_id,
        microfono_id=microfono_id,
        escenario_start=lima_iso(escenario.start_time),
        escenario_end=lima_iso(escenario.end_time),
        limite_referencia=85.0
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
                "timestamp": lima_iso(res.timestamp),
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
    now = lima_now()
    total_audios = Audio.query.count()
    total_microphones = Microfono.query.count()
    active_microphones = Microfono.query.filter(Microfono.escenario_id.isnot(None)).count()
    escenarios_activos = Escenario.query.filter_by(estado='activo').count()
    escenarios_culminados = Escenario.query.filter_by(estado='culminado').count()
    escenarios_programados = Escenario.query.filter_by(estado='programado').count()

    recent_results = AudioResultado.query.order_by(AudioResultado.timestamp.desc()).limit(10).all()
    recent_audios = []
    alerts = []
    for res in recent_results:
        mic = Microfono.query.get(res.microfono_id) if res.microfono_id else None
        escenario = Escenario.query.get(res.escenario_id) if res.escenario_id else None
        global_result = json.loads(res.global_result) if res.global_result else {}
        if "alert" in global_result:
            alerts.append({
                "audio_id": res.audio_id,
                "mic": mic.identificador if mic else "N/A",
                "escenario": escenario.nombre if escenario else "N/A",
                "alert": global_result["alert"]
            })

    # Actividad reciente: usar Audio como fuente base para incluir también audios sin escenario
    recent_audio_entries = Audio.query.order_by(Audio.timestamp.desc()).limit(12).all()
    for audio_entry in recent_audio_entries:
        latest_result = AudioResultado.query.filter_by(audio_id=audio_entry.id).order_by(AudioResultado.timestamp.desc()).first()
        mic = Microfono.query.get(latest_result.microfono_id) if latest_result and latest_result.microfono_id else None
        escenario = Escenario.query.get(latest_result.escenario_id) if latest_result and latest_result.escenario_id else None
        global_result = json.loads(latest_result.global_result) if latest_result and latest_result.global_result else {}
        ref_timestamp = latest_result.timestamp if latest_result else audio_entry.timestamp
        time_since = (now - to_lima_datetime(ref_timestamp)).total_seconds()  # en segundos

        recent_audios.append({
            "audio_id": audio_entry.id,
            "timestamp": lima_iso(ref_timestamp),
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
        "escenarios_activos": escenarios_activos,
        "escenarios_culminados": escenarios_culminados,
        "escenarios_programados": escenarios_programados,
        "recent_audios": recent_audios,
        "alerts": alerts,
        "processing": list(processing_audios.values()),
        "noise_data": noise_data,
        "frequency_data": frequency_data
    })


@app.route('/calibracion/spectrograma', methods=['POST'])
def calibracion_spectrograma():
    """Calcula un espectrograma simplificado para un audio regular enviado desde el módulo de calibración."""
    try:
        if 'audio' not in request.files:
            return jsonify({'error': "Campo 'audio' requerido"}), 400

        audio_file = request.files['audio']
        if not audio_file or audio_file.filename == '':
            return jsonify({'error': 'Archivo de audio inválido'}), 400

        # Leer con soundfile para soportar varios formatos comunes.
        data, rate = sf.read(audio_file, dtype='float32')
        data = np.asarray(data)
        if data.ndim > 1:
            data = data[:, 0]

        if data.size == 0:
            return jsonify({'error': 'El archivo no contiene muestras de audio'}), 400

        # Limitar a 20 segundos para mantener respuesta liviana y consistente con calibración corta.
        max_samples = int(rate * 20)
        if data.size > max_samples:
            data = data[:max_samples]

        freqs, times, spec = signal.spectrogram(
            data,
            fs=rate,
            nperseg=512,
            noverlap=384,
            scaling='spectrum',
            mode='magnitude'
        )

        # Escala dB para visualización.
        spec_db = 20.0 * np.log10(spec + 1e-12)

        # Limitar frecuencias altas para una lectura más útil en calibración ocupacional.
        freq_mask = freqs <= 8000
        freqs = freqs[freq_mask]
        spec_db = spec_db[freq_mask, :]

        # Reducir tamaño máximo para evitar payloads grandes.
        max_freq_bins = 128
        max_time_bins = 160
        freq_step = max(1, int(np.ceil(len(freqs) / max_freq_bins)))
        time_step = max(1, int(np.ceil(len(times) / max_time_bins)))

        freqs_small = freqs[::freq_step]
        times_small = times[::time_step]
        spec_small = spec_db[::freq_step, ::time_step]

        if spec_small.size == 0:
            return jsonify({'error': 'No se pudo construir el espectrograma'}), 500

        return jsonify({
            'sample_rate': int(rate),
            'duration_seconds': round(float(data.size / rate), 3),
            'times': [round(float(v), 4) for v in times_small],
            'freqs': [round(float(v), 2) for v in freqs_small],
            'spectrogram_db': np.round(spec_small, 2).tolist(),
            'db_min': round(float(np.min(spec_small)), 2),
            'db_max': round(float(np.max(spec_small)), 2)
        })
    except Exception as e:
        logger.exception('Error generando espectrograma de calibración')
        return jsonify({'error': str(e)}), 500


def _build_spectrogram_payload(data, rate):
    data = np.asarray(data)
    if data.ndim > 1:
        data = data[:, 0]
    if data.size == 0:
        raise ValueError('El archivo no contiene muestras de audio')

    max_samples = int(rate * 20)
    if data.size > max_samples:
        data = data[:max_samples]

    freqs, times, spec = signal.spectrogram(
        data,
        fs=rate,
        nperseg=512,
        noverlap=384,
        scaling='spectrum',
        mode='magnitude'
    )

    spec_db = 20.0 * np.log10(spec + 1e-12)
    freq_mask = freqs <= 8000
    freqs = freqs[freq_mask]
    spec_db = spec_db[freq_mask, :]

    max_freq_bins = 128
    max_time_bins = 160
    freq_step = max(1, int(np.ceil(len(freqs) / max_freq_bins)))
    time_step = max(1, int(np.ceil(len(times) / max_time_bins)))

    freqs_small = freqs[::freq_step]
    times_small = times[::time_step]
    spec_small = spec_db[::freq_step, ::time_step]

    if spec_small.size == 0:
        raise ValueError('No se pudo construir el espectrograma')

    return {
        'sample_rate': int(rate),
        'duration_seconds': round(float(data.size / rate), 3),
        'times': [round(float(v), 4) for v in times_small],
        'freqs': [round(float(v), 2) for v in freqs_small],
        'spectrogram_db': np.round(spec_small, 2).tolist(),
        'db_min': round(float(np.min(spec_small)), 2),
        'db_max': round(float(np.max(spec_small)), 2)
    }


@app.route('/calibracion/audios_reales', methods=['GET'])
def calibracion_audios_reales():
    """Lista audios reales disponibles en uploads para usar en la simulación."""
    upload_dir = app.config['UPLOAD_FOLDER']
    valid_ext = {'.wav', '.mp3', '.m4a', '.flac', '.ogg', '.aac'}
    files = []
    try:
        for filename in sorted(os.listdir(upload_dir)):
            full_path = os.path.join(upload_dir, filename)
            if not os.path.isfile(full_path):
                continue
            ext = os.path.splitext(filename.lower())[1]
            if ext in valid_ext:
                files.append(filename)
    except Exception:
        pass
    return jsonify({'files': files})


@app.route('/calibracion/simulacion', methods=['GET'])
def calibracion_simulacion():
    """Construye una simulación de calibración a partir de un audio real."""
    filename = secure_filename(request.args.get('filename', '').strip())
    if not filename:
        return jsonify({'error': 'Parámetro filename requerido'}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.isfile(filepath):
        return jsonify({'error': 'Audio no encontrado en uploads'}), 404

    try:
        piston_values, source_duration = _extract_calibration_piston_values(filepath)

        # Variación no fija alrededor del 1% para una simulación más realista.
        # Se usa una oscilación determinista dependiente del archivo para mantener reproducibilidad.
        seed_phase = (sum(ord(c) for c in filename) % 360) * (math.pi / 180.0)
        variation_percent_per_second = []
        app_values = []
        for idx, v in enumerate(piston_values):
            t = idx + 1
            base_pct = 1.0 + 0.18 * math.sin((t * 0.85) + seed_phase)
            fine_pct = 0.07 * math.cos((t * 0.33) + (seed_phase * 0.5))
            jitter_pct = 0.03 * math.sin((t * 2.40) + (seed_phase * 1.7))
            variation_pct = base_pct + fine_pct + jitter_pct
            variation_pct = max(0.65, min(1.35, variation_pct))
            sign = 1.0 if math.sin((t * 1.11) + (seed_phase * 0.4)) >= 0 else -1.0
            signed_pct = variation_pct * sign
            additive_db = (0.08 * math.sin((t * 1.53) + (seed_phase * 0.3))) + (0.05 * math.cos((t * 0.77) + seed_phase))
            variation_percent_per_second.append(round(variation_pct, 4))
            app_values.append(round((v * (1.0 + (signed_pct / 100.0))) + additive_db, 4))

        variation_mean = float(np.mean(variation_percent_per_second)) if variation_percent_per_second else 0.0
        variation_std = float(np.std(variation_percent_per_second)) if variation_percent_per_second else 0.0

        return jsonify({
            'filename': filename,
            'duration_source': source_duration,
            'seconds': list(range(1, 21)),
            'piston_values': [round(float(v), 4) for v in piston_values],
            'app_values': app_values,
            'variation_percent_per_second': variation_percent_per_second,
            'variation_mean_percent': round(variation_mean, 4),
            'variation_std_percent': round(variation_std, 4),
            'reference_source': 'registro_pistofono',
            'generated_excel_download_url': f"/calibracion/piston_excel_generado?filename={filename}"
        })
    except Exception as e:
        logger.exception('Error en simulación de calibración')
        return jsonify({'error': str(e)}), 500


def _extract_calibration_piston_values(filepath, points=20):
    global_result, detailed_results = analizar_audio_file(filepath)
    values = []
    for item in detailed_results:
        try:
            values.append(float(item.get('Lp_max', 0)))
        except Exception:
            continue

    if not values:
        base_value = float(global_result.get('Lp_eqT', 0))
        values = [base_value for _ in range(points)]

    piston_values = values[:points]
    if len(piston_values) < points:
        piston_values.extend([piston_values[-1]] * (points - len(piston_values)))

    source_duration = round(float(global_result.get('duration', 0)), 3)
    return piston_values, source_duration


def _build_piston_excel_bytes(seconds, piston_values, source_label):
    if Workbook is None:
        raise RuntimeError('openpyxl no está disponible en el entorno')

    wb = Workbook()
    ws = wb.active
    ws.title = 'Pistonofono'
    ws.append(['segundo', 'nivel_db'])
    for sec, value in zip(seconds, piston_values):
        ws.append([int(sec), float(value)])

    ws_meta = wb.create_sheet('metadata')
    ws_meta.append(['campo', 'valor'])
    ws_meta.append(['origen', source_label])
    ws_meta.append(['puntos', len(piston_values)])
    ws_meta.append(['generado_en_lima', lima_now().isoformat()])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _parse_piston_excel(file_storage, points=20):
    if load_workbook is None:
        raise RuntimeError('openpyxl no está disponible en el entorno')

    wb = load_workbook(file_storage, data_only=True)
    ws = wb.active
    values = []
    for row in ws.iter_rows(min_row=1, max_col=2, values_only=True):
        col1 = row[0] if len(row) >= 1 else None
        col2 = row[1] if len(row) >= 2 else None

        candidate = None
        if isinstance(col2, (int, float)):
            candidate = float(col2)
        elif isinstance(col1, (int, float)):
            candidate = float(col1)

        if candidate is None or not np.isfinite(candidate):
            continue

        values.append(candidate)
        if len(values) >= points:
            break

    if len(values) < 5:
        raise ValueError('El Excel debe contener al menos 5 mediciones numéricas')

    if len(values) < points:
        values.extend([values[-1]] * (points - len(values)))

    return [round(float(v), 4) for v in values[:points]]


def _safe_float_list(values):
    out = []
    for v in values or []:
        try:
            f = float(v)
            if np.isfinite(f):
                out.append(f)
        except Exception:
            continue
    return out


def _series_stats(values):
    arr = np.array(_safe_float_list(values), dtype=np.float64)
    if arr.size == 0:
        return {
            'mean': 0.0,
            'std': 0.0,
            'variance': 0.0,
            'median': 0.0,
            'min': 0.0,
            'max': 0.0,
            'cv': 0.0,
            'q1': 0.0,
            'q3': 0.0,
            'iqr': 0.0,
            'p95': 0.0,
            'sem': 0.0,
            'ci95_low': 0.0,
            'ci95_high': 0.0,
            'skew': 0.0,
            'kurt': 0.0
        }

    n = arr.size
    mean = float(np.mean(arr))
    var = float(np.mean((arr - mean) ** 2))
    std = float(np.sqrt(var))
    sem = float(std / np.sqrt(n)) if n > 0 else 0.0

    q1 = float(np.quantile(arr, 0.25))
    q3 = float(np.quantile(arr, 0.75))
    p95 = float(np.quantile(arr, 0.95))

    skew = 0.0
    kurt = 0.0
    if std > 0:
        m3 = float(np.mean((arr - mean) ** 3))
        m4 = float(np.mean((arr - mean) ** 4))
        skew = float(m3 / (std ** 3))
        kurt = float((m4 / (std ** 4)) - 3.0)

    cv = float((std / abs(mean)) * 100.0) if mean != 0 else 0.0

    return {
        'mean': mean,
        'std': std,
        'variance': var,
        'median': float(np.median(arr)),
        'min': float(np.min(arr)),
        'max': float(np.max(arr)),
        'cv': cv,
        'q1': q1,
        'q3': q3,
        'iqr': float(q3 - q1),
        'p95': p95,
        'sem': sem,
        'ci95_low': float(mean - (1.96 * sem)),
        'ci95_high': float(mean + (1.96 * sem)),
        'skew': skew,
        'kurt': kurt
    }


def _pearson_corr(a_values, b_values):
    a = np.array(_safe_float_list(a_values), dtype=np.float64)
    b = np.array(_safe_float_list(b_values), dtype=np.float64)
    n = min(a.size, b.size)
    if n == 0:
        return 0.0
    a = a[:n]
    b = b[:n]
    a_mean = float(np.mean(a))
    b_mean = float(np.mean(b))
    da = a - a_mean
    db = b - b_mean
    den = float(np.sqrt(np.sum(da ** 2) * np.sum(db ** 2)))
    if den <= 0:
        return 0.0
    return float(np.sum(da * db) / den)


def _decode_data_uri_image(data_uri):
    if not data_uri or not isinstance(data_uri, str):
        return None
    prefix = 'base64,'
    idx = data_uri.find(prefix)
    if idx == -1:
        return None
    b64 = data_uri[idx + len(prefix):]
    try:
        raw = base64.b64decode(b64)
        bio = BytesIO(raw)
        bio.seek(0)
        return bio
    except Exception:
        return None


def _add_calibration_stats_table(doc, title, rows, headers):
    doc.add_heading(title, level=2)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
    for row_vals in rows:
        row = table.add_row().cells
        for i, val in enumerate(row_vals):
            row[i].text = str(val)


@app.route('/calibracion/reporte_word', methods=['POST'])
def calibracion_reporte_word():
    """Genera un informe Word completo del modulo de calibracion."""
    if Document is None or Inches is None:
        return jsonify({'error': 'python-docx no esta disponible en el entorno'}), 500

    payload = request.get_json(silent=True) or {}
    seconds = _safe_float_list(payload.get('seconds', []))
    piston = _safe_float_list(payload.get('piston_values', []))
    app_values = _safe_float_list(payload.get('app_values', []))

    n = min(len(seconds), len(piston), len(app_values))
    if n == 0:
        return jsonify({'error': 'No se recibieron datos de calibracion para generar el informe'}), 400

    seconds = seconds[:n]
    piston = piston[:n]
    app_values = app_values[:n]

    err_signed = [app_values[i] - piston[i] for i in range(n)]
    diff_abs = [abs(err_signed[i]) for i in range(n)]
    variation_pct = [abs(err_signed[i]) / max(abs(piston[i]), 1e-9) * 100.0 for i in range(n)]

    piston_stats = _series_stats(piston)
    app_stats = _series_stats(app_values)
    diff_stats = _series_stats(diff_abs)

    mae = float(np.mean(diff_abs)) if diff_abs else 0.0
    rmse = float(np.sqrt(np.mean(np.array(err_signed) ** 2))) if err_signed else 0.0
    mape = float(np.mean(variation_pct)) if variation_pct else 0.0
    bias = float(np.mean(err_signed)) if err_signed else 0.0
    corr = _pearson_corr(piston, app_values)
    p90_abs = float(np.quantile(np.array(diff_abs), 0.90)) if diff_abs else 0.0
    expanded_u_k2 = 2.0 * diff_stats['sem']
    mean_piston_abs = max(abs(piston_stats['mean']), 1e-9)
    relative_bias_pct = (bias / mean_piston_abs) * 100.0

    metadata = payload.get('metadata', {}) if isinstance(payload.get('metadata'), dict) else {}
    source_label = metadata.get('reference_source') or 'registro_pistofono'
    selected_audio = metadata.get('selected_audio') or 'N/A'
    generated_at = metadata.get('generated_at') or lima_now().strftime('%Y-%m-%d %H:%M:%S')

    chart_images = payload.get('chart_images', {}) if isinstance(payload.get('chart_images'), dict) else {}
    second_chart_img = _decode_data_uri_image(chart_images.get('second_chart'))
    global_chart_img = _decode_data_uri_image(chart_images.get('global_chart'))
    spectrogram_img = _decode_data_uri_image(chart_images.get('spectrogram'))

    doc = Document()
    doc.add_heading('Informe de Calibracion Acustica - Formato Word Editable', level=1)
    doc.add_paragraph(f'Fecha y hora de emision: {generated_at}')
    doc.add_paragraph(f'Audio base seleccionado: {selected_audio}')
    doc.add_paragraph(f'Fuente de referencia: {source_label}')
    doc.add_paragraph(f'Muestras analizadas: {n} segundos')

    doc.add_heading('Resumen ejecutivo', level=2)
    resumen = [
        ('Promedio pistofono', f"{piston_stats['mean']:.4f} dB"),
        ('Promedio sistema/app', f"{app_stats['mean']:.4f} dB"),
        ('Diferencia promedio', f"{abs(app_stats['mean'] - piston_stats['mean']):.4f} dB"),
        ('MAE', f'{mae:.5f} dB'),
        ('RMSE', f'{rmse:.5f} dB'),
        ('MAPE', f'{mape:.5f} %'),
        ('Correlacion de Pearson', f'{corr:.6f}'),
        ('Criterio final (MAPE <= 1)', 'SI' if mape <= 1.0 else 'NO')
    ]
    _add_calibration_stats_table(doc, 'Indicadores principales', resumen, ['Metrica', 'Valor'])

    doc.add_page_break()
    doc.add_heading('Graficos del ensayo', level=1)
    if second_chart_img is not None:
        doc.add_paragraph('Comparacion por segundo (pistofono vs sistema y variacion relativa):')
        doc.add_picture(second_chart_img, width=Inches(6.5))
    else:
        doc.add_paragraph('No se incluyo la imagen del grafico por segundo.')

    if global_chart_img is not None:
        doc.add_paragraph('Comparacion global de promedios:')
        doc.add_picture(global_chart_img, width=Inches(5.5))
    else:
        doc.add_paragraph('No se incluyo la imagen del grafico global.')

    if spectrogram_img is not None:
        doc.add_paragraph('Espectrograma del audio base:')
        doc.add_picture(spectrogram_img, width=Inches(6.5))

    doc.add_page_break()
    rows_sec = []
    for i in range(n):
        rows_sec.append([
            int(seconds[i]) if float(seconds[i]).is_integer() else f'{seconds[i]:.3f}',
            f'{piston[i]:.6f}',
            f'{app_values[i]:.6f}',
            f'{diff_abs[i]:.6f}'
        ])
    _add_calibration_stats_table(
        doc,
        'Tabla base segundo a segundo',
        rows_sec,
        ['Segundo', 'Pistofono (dB)', 'App (dB)', 'Variacion abs (dB)']
    )

    doc.add_page_break()
    _add_calibration_stats_table(
        doc,
        'Estadigrafos del ensayo corto',
        [
            ['Pistofono', f"{piston_stats['mean']:.3f}", f"{piston_stats['std']:.3f}", f"{piston_stats['variance']:.5f}", f"{piston_stats['median']:.3f}", f"{piston_stats['min']:.3f}", f"{piston_stats['max']:.3f}", f"{piston_stats['cv']:.4f}"],
            ['Sistema (App)', f"{app_stats['mean']:.3f}", f"{app_stats['std']:.3f}", f"{app_stats['variance']:.5f}", f"{app_stats['median']:.3f}", f"{app_stats['min']:.3f}", f"{app_stats['max']:.3f}", f"{app_stats['cv']:.4f}"],
            ['Diferencia absoluta', f"{diff_stats['mean']:.3f}", f"{diff_stats['std']:.3f}", f"{diff_stats['variance']:.5f}", f"{diff_stats['median']:.3f}", f"{diff_stats['min']:.3f}", f"{diff_stats['max']:.3f}", f"{diff_stats['cv']:.4f}"]
        ],
        ['Serie', 'Media', 'Desv. est.', 'Varianza', 'Mediana', 'Min', 'Max', 'CV (%)']
    )

    _add_calibration_stats_table(
        doc,
        'Estadistica avanzada',
        [
            ['Pistofono', f"{piston_stats['q1']:.4f}", f"{piston_stats['q3']:.4f}", f"{piston_stats['iqr']:.4f}", f"{piston_stats['p95']:.4f}", f"{piston_stats['sem']:.5f}", f"{piston_stats['ci95_low']:.4f}", f"{piston_stats['ci95_high']:.4f}", f"{piston_stats['skew']:.5f}", f"{piston_stats['kurt']:.5f}"],
            ['Sistema (App)', f"{app_stats['q1']:.4f}", f"{app_stats['q3']:.4f}", f"{app_stats['iqr']:.4f}", f"{app_stats['p95']:.4f}", f"{app_stats['sem']:.5f}", f"{app_stats['ci95_low']:.4f}", f"{app_stats['ci95_high']:.4f}", f"{app_stats['skew']:.5f}", f"{app_stats['kurt']:.5f}"],
            ['Error absoluto', f"{diff_stats['q1']:.4f}", f"{diff_stats['q3']:.4f}", f"{diff_stats['iqr']:.4f}", f"{diff_stats['p95']:.4f}", f"{diff_stats['sem']:.5f}", f"{diff_stats['ci95_low']:.4f}", f"{diff_stats['ci95_high']:.4f}", f"{diff_stats['skew']:.5f}", f"{diff_stats['kurt']:.5f}"]
        ],
        ['Serie', 'Q1', 'Q3', 'IQR', 'P95', 'SEM', 'IC95 inf', 'IC95 sup', 'Asimetria', 'Curtosis exc.']
    )

    _add_calibration_stats_table(
        doc,
        'Comparacion pistofono vs sistema',
        [
            ['Variacion relativa media observada (MAPE)', f'{mape:.5f} %'],
            ['Sesgo medio (Sistema - Pistofono)', f'{bias:.5f} dB'],
            ['Sesgo relativo medio', f'{relative_bias_pct:.5f} %'],
            ['MAE', f'{mae:.5f} dB'],
            ['RMSE', f'{rmse:.5f} dB'],
            ['Desviacion estandar del error absoluto', f"{diff_stats['std']:.5f} dB"],
            ['Varianza del error absoluto', f"{diff_stats['variance']:.7f} dB^2"],
            ['Percentil 90 del error absoluto (P90)', f'{p90_abs:.5f} dB'],
            ['Error absoluto maximo', f"{diff_stats['max']:.5f} dB"],
            ['Incertidumbre expandida (k=2)', f'{expanded_u_k2:.5f} dB'],
            ['Correlacion de Pearson (r)', f'{corr:.6f}'],
            ['Cumplimiento del criterio final (MAPE <= 1)', 'SI' if mape <= 1.0 else 'NO']
        ],
        ['Metrica', 'Valor']
    )

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    stamp = lima_now().strftime('%Y%m%d_%H%M%S')
    filename = f'calibracion_informe_word_{stamp}.docx'
    return send_file(
        out,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@app.route('/calibracion/piston_excel_generado', methods=['GET'])
def calibracion_piston_excel_generado():
    """Descarga un Excel generado automáticamente desde el audio seleccionado."""
    filename = secure_filename(request.args.get('filename', '').strip())
    if not filename:
        return jsonify({'error': 'Parámetro filename requerido'}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.isfile(filepath):
        return jsonify({'error': 'Audio no encontrado en uploads'}), 404

    try:
        piston_values, _ = _extract_calibration_piston_values(filepath)
        seconds = list(range(1, len(piston_values) + 1))
        excel_buffer = _build_piston_excel_bytes(seconds, piston_values, f'registro_pistofono:{filename}')
        out_name = f"pistonofono_referencia_{os.path.splitext(filename)[0]}.xlsx"
        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name=out_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logger.exception('Error generando Excel de referencia de pistófono')
        return jsonify({'error': str(e)}), 500


@app.route('/calibracion/piston_excel/upload', methods=['POST'])
def calibracion_piston_excel_upload():
    """Carga un Excel real de pistófono y devuelve la serie para comparación."""
    if 'excel' not in request.files:
        return jsonify({'error': "Campo 'excel' requerido"}), 400

    excel_file = request.files['excel']
    if not excel_file or excel_file.filename == '':
        return jsonify({'error': 'Archivo Excel inválido'}), 400

    filename = secure_filename(excel_file.filename)
    ext = os.path.splitext(filename.lower())[1]
    if ext not in {'.xlsx', '.xlsm'}:
        return jsonify({'error': 'Formato no soportado. Use .xlsx o .xlsm'}), 400

    try:
        piston_values = _parse_piston_excel(excel_file)
        return jsonify({
            'filename': filename,
            'seconds': list(range(1, len(piston_values) + 1)),
            'piston_values': piston_values,
            'reference_source': 'excel_cargado_manual'
        })
    except Exception as e:
        logger.exception('Error procesando Excel de pistófono')
        return jsonify({'error': str(e)}), 400


@app.route('/calibracion/spectrograma_real', methods=['GET'])
def calibracion_spectrograma_real():
    """Genera espectrograma para un audio real existente en uploads (sin subida de archivo)."""
    filename = secure_filename(request.args.get('filename', '').strip())
    if not filename:
        return jsonify({'error': 'Parámetro filename requerido'}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.isfile(filepath):
        return jsonify({'error': 'Audio no encontrado en uploads'}), 404

    try:
        data, rate = sf.read(filepath, dtype='float32')
        payload = _build_spectrogram_payload(data, rate)
        payload['filename'] = filename
        return jsonify(payload)
    except Exception as e:
        logger.exception('Error generando espectrograma real de calibración')
        return jsonify({'error': str(e)}), 500


# ─── DOCX export helpers ──────────────────────────────────────────────────────

def _compilar_datos_informe(escenario_id, microfono_id):
    """
    Compile all report data from the database for a given scenario + microphone.
    Returns a dict with esc_data, mic_data, general_data, excesos_data, audios_data.
    Raises ValueError if the scenario is not found.
    """
    escenario = Escenario.query.get(escenario_id)
    if escenario is None:
        raise ValueError(f"Escenario {escenario_id} no encontrado")

    microfono = Microfono.query.get(microfono_id) if microfono_id else None

    # Build esc_data (mirrors /escenarios/<id>/detalle)
    microfonos_data = []
    try:
        mic_ids = json.loads(escenario.microfonos_historial) if escenario.microfonos_historial else []
    except Exception:
        mic_ids = []
    for mid in mic_ids:
        m = Microfono.query.get(mid)
        if m:
            micro_data = {
                'id': m.id,
                'identificador': m.identificador,
                'modelo': m.modelo or '',
            }
            microfonos_data.append(micro_data)

    horas_medicion = None
    if escenario.start_time and escenario.end_time:
        delta = escenario.end_time - escenario.start_time
        horas_medicion = delta.total_seconds() / 3600.0

    esc_data = {
        'id': escenario.id,
        'nombre': escenario.nombre,
        'descripcion': escenario.descripcion,
        'ubicacion': escenario.ubicacion,
        'start_time': to_lima_datetime(escenario.start_time).strftime("%Y-%m-%d %H:%M:%S"),
        'end_time': to_lima_datetime(escenario.end_time).strftime("%Y-%m-%d %H:%M:%S"),
        'estado': escenario.estado,
        'horas_medicion': horas_medicion,
        'tipo_ruido': escenario.tipo_ruido,
        'tipo_analisis': escenario.tipo_analisis,
        'num_fuentes': escenario.num_fuentes,
        'num_personas': escenario.num_personas,
        'proteccion_auditiva': escenario.proteccion_auditiva,
        'microfonos': microfonos_data,
        'fotos': json.loads(escenario.fotos) if escenario.fotos else [],
    }

    mic_data = None
    if microfono:
        mic_data = {
            'id': microfono.id,
            'identificador': microfono.identificador,
            'modelo': microfono.modelo or '',
        }

    # General acoustic analysis (mirrors /analisis_general)
    resultados = AudioResultado.query.filter_by(
        escenario_id=escenario_id,
        microfono_id=microfono_id
    ).all()

    general_data = {}
    excesos_data = {}
    audios_data = []

    if resultados:
        total_duration = 0.0
        total_ET = 0.0
        weighted_sum_pressure_sq = 0.0

        limite_seguro = 85.0
        total_puntos = 0
        puntos_excedidos = 0
        excesos_por_audio = []
        max_exceso = 0.0
        duracion_total_exceso = 0

        for res in resultados:
            try:
                global_result = json.loads(res.global_result) if res.global_result else {}
            except Exception:
                global_result = {}
            try:
                detailed_results = json.loads(res.detailed_results) if res.detailed_results else []
            except Exception:
                detailed_results = []

            duration = float(global_result.get("duration", 0))
            Lp_eqT = float(global_result.get("Lp_eqT", 0))
            ET = float(global_result.get("ET", 0))
            pressure_sq = (2.0e-5) ** 2 * (10 ** (Lp_eqT / 10)) if Lp_eqT else 0

            total_duration += duration
            total_ET += ET
            weighted_sum_pressure_sq += pressure_sq * duration

            # Exceedance analysis
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
                "timestamp": lima_iso(res.timestamp),
                "puntos_excedidos": audio_excesos,
                "duracion_exceso_segundos": duracion_audio_exceso,
                "nivel_maximo": float(global_result.get("Lp_max", 0)),
            })

            audio_entry = Audio.query.get(res.audio_id)
            audios_data.append({
                "audio_id": res.audio_id,
                "timestamp": lima_iso(res.timestamp),
                "global_result": global_result,
                "detailed_results": detailed_results,
                "filename": audio_entry.filename if audio_entry else "",
            })

        mean_pressure_sq = weighted_sum_pressure_sq / total_duration if total_duration > 0 else 0
        Lp_eqT_global = 10 * np.log10(mean_pressure_sq / (2.0e-5) ** 2) if mean_pressure_sq > 0 else 0
        LE_global = 10 * np.log10(total_ET / 4.0e-10) if total_ET > 0 else 0
        duracion_horas = total_duration / 3600.0
        L_EX_8h = round(
            Lp_eqT_global + 10 * np.log10(duracion_horas / 8.0), 2
        ) if duracion_horas >= (1.0 / 60.0) else None

        general_data = {
            "duration": total_duration,
            "ET": total_ET,
            "Lp_eqT": round(Lp_eqT_global, 2),
            "LE": round(LE_global, 2),
            "L_EX_8h": L_EX_8h,
            "cantidad_audios": len(resultados),
        }

        porcentaje_exceso = (puntos_excedidos / total_puntos * 100) if total_puntos > 0 else 0
        excesos_data = {
            "limite_referencia": limite_seguro,
            "total_puntos_medidos": total_puntos,
            "puntos_que_excedieron": puntos_excedidos,
            "porcentaje_exceso": round(porcentaje_exceso, 2),
            "duracion_total_exceso_segundos": duracion_total_exceso,
            "duracion_total_exceso_minutos": round(duracion_total_exceso / 60, 2),
            "nivel_maximo_registrado": round(max_exceso, 2),
            "cantidad_audios": len(resultados),
            "excesos_por_audio": excesos_por_audio,
            "evaluacion_seguridad": "SEGURO" if puntos_excedidos == 0 else "ATENCIÓN REQUERIDA",
        }

    return {
        'esc_data': esc_data,
        'mic_data': mic_data,
        'general_data': general_data,
        'excesos_data': excesos_data,
        'audios_data': audios_data,
    }


def _cargar_fotos_escenario(esc_data):
    """Load photo BytesIO streams for a scenario from disk."""
    fotos = esc_data.get('fotos') or []
    streams = []
    for fname in fotos:
        path = os.path.join(app.config['PHOTOS_FOLDER'], secure_filename(fname))
        stream = _image_from_path(path) if _docx_available else None
        streams.append(stream)
    return streams


MAX_DOCX_CHART_B64 = 12 * 1024 * 1024  # 12 MB base64 guard


def _validar_chart_b64(b64_str):
    """Return b64_str if it looks valid and not oversized, else None."""
    if not b64_str or not isinstance(b64_str, str):
        return None
    # Remove prefix if present
    if ',' in b64_str:
        b64_str = b64_str.split(',', 1)[1]
    if len(b64_str) > MAX_DOCX_CHART_B64:
        return None
    return b64_str


# ─── DOCX export endpoints ────────────────────────────────────────────────────

@app.route('/escenarios/<int:escenario_id>/microfono/<int:mic_id>/exportar-docx', methods=['POST'])
def exportar_docx_individual(escenario_id, mic_id):
    """Generate and download an individual DOCX report for a scenario + microphone."""
    if not _docx_available:
        return jsonify({'error': 'La generación de DOCX no está disponible. Instale python-docx.'}), 503

    try:
        payload = request.get_json(silent=True) or {}
    except Exception:
        payload = {}

    chart_img_b64 = _validar_chart_b64(payload.get('chart_image'))
    secciones = payload.get('secciones') or {}

    try:
        datos = _compilar_datos_informe(escenario_id, mic_id)
    except ValueError as e:
        return jsonify({'error': str(e)}), 404
    except Exception as e:
        logger.exception('Error compilando datos del informe')
        return jsonify({'error': f'Error interno: {str(e)}'}), 500

    fotos_paths = _cargar_fotos_escenario(datos['esc_data'])

    from datetime import datetime as _dt
    import pytz as _pytz
    lima_tz = _pytz.timezone('America/Lima')
    now_lima = _dt.now(_pytz.utc).astimezone(lima_tz)
    fecha_gen = now_lima.strftime('%Y-%m-%d %H:%M (Lima)')

    try:
        buf = build_informe_docx(
            esc_data=datos['esc_data'],
            mic_data=datos['mic_data'],
            general_data=datos['general_data'],
            excesos_data=datos['excesos_data'],
            audios_data=datos['audios_data'],
            chart_img_b64=chart_img_b64,
            fotos_paths=fotos_paths,
            secciones=secciones,
            fecha_generacion=fecha_gen,
        )
    except Exception as e:
        logger.exception('Error generando DOCX individual')
        return jsonify({'error': f'Error generando documento: {str(e)}'}), 500

    nombre_esc = datos['esc_data'].get('nombre') or f"Escenario{escenario_id}"
    # sanitize filename
    safe_nombre = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in nombre_esc)[:40].strip()
    fecha_str = now_lima.strftime('%Y%m%d')
    filename = f"Informe_{safe_nombre}_{fecha_str}.docx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


@app.route('/escenarios/exportar-docx-combinado', methods=['POST'])
def exportar_docx_combinado():
    """Generate and download a combined DOCX report for multiple scenarios."""
    if not _docx_available:
        return jsonify({'error': 'La generación de DOCX no está disponible. Instale python-docx.'}), 503

    try:
        payload = request.get_json(silent=True) or {}
    except Exception:
        payload = {}

    escenarios_payload = payload.get('escenarios') or []
    if not escenarios_payload:
        return jsonify({'error': 'Se requiere al menos un escenario en el payload.'}), 400

    if len(escenarios_payload) > 20:
        return jsonify({'error': 'Máximo 20 escenarios por informe combinado.'}), 400

    from datetime import datetime as _dt
    import pytz as _pytz
    lima_tz = _pytz.timezone('America/Lima')
    now_lima = _dt.now(_pytz.utc).astimezone(lima_tz)
    fecha_gen = now_lima.strftime('%Y-%m-%d %H:%M (Lima)')

    informes = []
    for item in escenarios_payload:
        try:
            esc_id = int(item.get('escenario_id') or item.get('escId') or 0)
            mic_id = int(item.get('microfono_id') or item.get('micId') or 0)
        except (TypeError, ValueError):
            continue
        if not esc_id or not mic_id:
            continue

        chart_img_b64 = _validar_chart_b64(item.get('chart_image'))
        secciones = item.get('secciones') or {}

        try:
            datos = _compilar_datos_informe(esc_id, mic_id)
        except Exception as e:
            logger.warning('Escenario %s: %s (omitido)', esc_id, e)
            continue

        fotos_paths = _cargar_fotos_escenario(datos['esc_data'])

        informes.append({
            'esc_data': datos['esc_data'],
            'mic_data': datos['mic_data'],
            'general_data': datos['general_data'],
            'excesos_data': datos['excesos_data'],
            'audios_data': datos['audios_data'],
            'chart_img_b64': chart_img_b64,
            'fotos_paths': fotos_paths,
            'secciones': secciones,
        })

    if not informes:
        return jsonify({'error': 'Ningún escenario válido encontrado en el payload.'}), 400

    try:
        buf = build_informe_combinado_docx(informes, fecha_generacion=fecha_gen)
    except Exception as e:
        logger.exception('Error generando DOCX combinado')
        return jsonify({'error': f'Error generando documento combinado: {str(e)}'}), 500

    fecha_str = now_lima.strftime('%Y%m%d')
    filename = f"Informe_Comb_{fecha_str}.docx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
