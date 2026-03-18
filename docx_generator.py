"""
docx_generator.py – Generación de informes Word (.docx) para el análisis acústico.

Formato:
  - Página A4 (210 × 297 mm)
  - Márgenes: izquierda 15 mm, derecha 15 mm, superior 15 mm, inferior 20 mm
  - Ancho útil: 180 mm
"""

import base64
import datetime
import os
from io import BytesIO

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

# ─── Constants ────────────────────────────────────────────────────────────────
PAGE_WIDTH_MM = 210
MARGIN_LEFT_MM = 15
MARGIN_RIGHT_MM = 15
MARGIN_TOP_MM = 15
MARGIN_BOTTOM_MM = 20
USABLE_WIDTH_MM = PAGE_WIDTH_MM - MARGIN_LEFT_MM - MARGIN_RIGHT_MM  # 180 mm

LIMITE_SEGURO_DB = 85.0   # Límite máximo permisible de referencia (dB)

MAX_PHOTO_BYTES = 10 * 1024 * 1024  # 10 MB guard
MAX_CHART_BYTES = 8 * 1024 * 1024   # 8 MB guard

# ─── XML/table helpers ────────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color):
    """Set cell background colour (e.g. '1e293b')."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def _set_cell_borders(cell):
    """Add thin borders to a cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'CBD5E1')
        borders.append(border)
    tcPr.append(borders)


def _set_table_no_split(table):
    """Prevent table rows from splitting across pages (best-effort)."""
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement('w:cantSplit')
        cant_split.set(qn('w:val'), '1')
        trPr.append(cant_split)


def _col_widths_twips(fractions, total_mm=USABLE_WIDTH_MM):
    """Return list of column widths in twips given fractional sizes."""
    total_twips = int(total_mm * 56.6929)  # 1 mm ≈ 56.6929 twips
    return [int(f * total_twips) for f in fractions]


def _set_col_widths(table, widths_twips):
    tbl = table._tbl
    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is None:
        tblGrid = OxmlElement('w:tblGrid')
        tbl.insert(0, tblGrid)
    else:
        for child in list(tblGrid):
            tblGrid.remove(child)
    for w in widths_twips:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(w))
        tblGrid.append(gridCol)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.append(tcW)
            w = widths_twips[idx] if idx < len(widths_twips) else widths_twips[-1]
            tcW.set(qn('w:w'), str(w))
            tcW.set(qn('w:type'), 'dxa')


def _add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


def _safe_b64_to_image(b64_str, max_bytes=MAX_PHOTO_BYTES):
    """Decode base64 string to BytesIO, returns None on error/oversized."""
    if not b64_str:
        return None
    try:
        # Strip data URI prefix if present
        if ',' in b64_str:
            b64_str = b64_str.split(',', 1)[1]
        raw = base64.b64decode(b64_str)
        if len(raw) > max_bytes:
            return None
        return BytesIO(raw)
    except Exception:
        return None


def _image_from_path(path, max_bytes=MAX_PHOTO_BYTES):
    """Read image file and return BytesIO, returns None on error/oversized."""
    try:
        if not os.path.isfile(path):
            return None
        size = os.path.getsize(path)
        if size > max_bytes:
            return None
        with open(path, 'rb') as f:
            return BytesIO(f.read())
    except Exception:
        return None


# ─── Document helpers ─────────────────────────────────────────────────────────

def _new_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(PAGE_WIDTH_MM)
    section.page_height = Mm(297)
    section.left_margin = Mm(MARGIN_LEFT_MM)
    section.right_margin = Mm(MARGIN_RIGHT_MM)
    section.top_margin = Mm(MARGIN_TOP_MM)
    section.bottom_margin = Mm(MARGIN_BOTTOM_MM)
    # Remove default empty paragraph
    for para in doc.paragraphs:
        p = para._element
        p.getparent().remove(p)
    return doc


def _heading(doc, text, level=1, color='1e293b'):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    size = {1: 16, 2: 13, 3: 11}.get(level, 11)
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(color)
    return p


def _para(doc, text='', size=10, bold=False, color='333333', space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    return p


def _add_kv_table(doc, rows, col_widths=None):
    """
    Add a two-column key-value table.
    rows = list of (key, value) tuples.
    """
    if col_widths is None:
        col_widths = [0.3, 0.7]
    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    widths = _col_widths_twips(col_widths)
    _set_col_widths(table, widths)
    for i, (k, v) in enumerate(rows):
        cell_k = table.rows[i].cells[0]
        cell_v = table.rows[i].cells[1]
        _set_cell_bg(cell_k, 'F1F5F9')
        _set_cell_borders(cell_k)
        _set_cell_borders(cell_v)
        cell_k.paragraphs[0].clear()
        run_k = cell_k.paragraphs[0].add_run(str(k))
        run_k.font.size = Pt(9)
        run_k.font.bold = True
        run_k.font.color.rgb = RGBColor.from_string('1e293b')
        cell_v.paragraphs[0].clear()
        run_v = cell_v.paragraphs[0].add_run(str(v))
        run_v.font.size = Pt(9)
        run_v.font.color.rgb = RGBColor.from_string('1e293b')
    doc.add_paragraph()
    return table


def _add_data_table(doc, headers, data_rows, col_widths=None):
    """
    Add a data table with a styled header row.
    headers = list of strings
    data_rows = list of lists of strings
    """
    n_cols = len(headers)
    if col_widths is None:
        col_widths = [1 / n_cols] * n_cols
    total_rows = 1 + len(data_rows)
    table = doc.add_table(rows=total_rows, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    widths = _col_widths_twips(col_widths)
    _set_col_widths(table, widths)
    # Header row
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        _set_cell_bg(cell, '1E293B')
        _set_cell_borders(cell)
        cell.paragraphs[0].clear()
        run = cell.paragraphs[0].add_run(str(h))
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Data rows
    for i, row_data in enumerate(data_rows):
        bg = 'F8FAFC' if i % 2 == 0 else 'FFFFFF'
        for j, val in enumerate(row_data):
            if j >= n_cols:
                break
            cell = table.rows[i + 1].cells[j]
            _set_cell_bg(cell, bg)
            _set_cell_borders(cell)
            cell.paragraphs[0].clear()
            run = cell.paragraphs[0].add_run(str(val) if val is not None else '—')
            run.font.size = Pt(8.5)
            run.font.color.rgb = RGBColor.from_string('1e293b')
    _set_table_no_split(table)
    doc.add_paragraph()
    return table


# ─── Section builders ─────────────────────────────────────────────────────────

def _sec_portada(doc, esc_data, mic_data, fecha_generacion):
    """Cover page."""
    doc.add_paragraph()
    doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run('INFORME TÉCNICO')
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string('64748B')

    p_main = doc.add_paragraph()
    p_main.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_main.add_run('Análisis de Exposición al Ruido Ocupacional')
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string('0F172A')

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_sub.add_run('Monitoreo Acústico de Ambiente Laboral')
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor.from_string('475569')

    doc.add_paragraph()
    doc.add_paragraph()

    # Scenario info block
    p_esc = doc.add_paragraph()
    p_esc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    nombre = esc_data.get('nombre') or f"Escenario #{esc_data.get('id', '')}"
    run = p_esc.add_run(nombre)
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string('1e293b')

    if esc_data.get('ubicacion'):
        p_ub = doc.add_paragraph()
        p_ub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_ub.add_run(f"Ubicación: {esc_data['ubicacion']}")
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor.from_string('475569')

    # Dates
    start = esc_data.get('start_time', '')
    end = esc_data.get('end_time', '')
    if start or end:
        p_dates = doc.add_paragraph()
        p_dates.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dates_str = ''
        if start and end:
            dates_str = f"Período: {start}  —  {end}"
        elif start:
            dates_str = f"Inicio: {start}"
        run = p_dates.add_run(dates_str)
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor.from_string('64748B')

    if mic_data:
        p_mic = doc.add_paragraph()
        p_mic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_mic.add_run(f"Micrófono: {mic_data.get('identificador', '')} · {mic_data.get('modelo', '')}")
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor.from_string('475569')

    doc.add_paragraph()
    doc.add_paragraph()

    p_gen = doc.add_paragraph()
    p_gen.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_gen.add_run(f"Generado: {fecha_generacion}")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string('94A3B8')

    _add_page_break(doc)


def _sec_escenario(doc, esc_data, mic_data):
    """Section I – Scenario information."""
    _heading(doc, 'I.  Información del Escenario')
    horas = esc_data.get('horas_medicion')
    horas_str = f"{horas:.2f}" if horas is not None else '—'
    mics = esc_data.get('microfonos') or []
    mic_str = ', '.join(m.get('identificador', '') for m in mics) if mics else (
        mic_data.get('identificador', '—') if mic_data else '—'
    )
    rows = [
        ('Nombre del Escenario', esc_data.get('nombre') or '—'),
        ('Ubicación', esc_data.get('ubicacion') or '—'),
        ('Descripción', esc_data.get('descripcion') or '—'),
        ('Inicio de Medición', esc_data.get('start_time') or '—'),
        ('Fin de Medición', esc_data.get('end_time') or '—'),
        ('Duración Total (h)', horas_str),
        ('Tipo de Ruido', esc_data.get('tipo_ruido') or '—'),
        ('Tipo de Análisis', esc_data.get('tipo_analisis') or '—'),
        ('N° de Fuentes', str(esc_data.get('num_fuentes') if esc_data.get('num_fuentes') is not None else '—')),
        ('N° de Personas', str(esc_data.get('num_personas') if esc_data.get('num_personas') is not None else '—')),
        ('Protección Auditiva', esc_data.get('proteccion_auditiva') or '—'),
        ('Micrófonos Asignados', mic_str),
        ('Estado', esc_data.get('estado') or '—'),
    ]
    _add_kv_table(doc, rows)


def _sec_fotos(doc, fotos_paths):
    """Section II – Site photographs."""
    _heading(doc, 'II.  Fotografías del Lugar de Medición')
    if not fotos_paths:
        _para(doc, '[Sin fotografías registradas para este escenario]', size=9, color='94A3B8')
        doc.add_paragraph()
        return

    # Insert photos two per row, scaled to half the usable width each
    photo_width = Mm(USABLE_WIDTH_MM / 2 - 3)  # leave a small gap
    loaded = 0
    for i in range(0, len(fotos_paths), 2):
        pair = fotos_paths[i:i + 2]
        table = doc.add_table(rows=1, cols=len(pair))
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        col_w = 1 / len(pair)
        widths = _col_widths_twips([col_w] * len(pair))
        _set_col_widths(table, widths)
        for j, img_stream in enumerate(pair):
            cell = table.rows[0].cells[j]
            cell.paragraphs[0].clear()
            try:
                cell.paragraphs[0].add_run().add_picture(img_stream, width=photo_width)
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                loaded += 1
            except Exception:
                run = cell.paragraphs[0].add_run('[Foto no disponible]')
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor.from_string('94A3B8')
        doc.add_paragraph()

    if loaded == 0:
        _para(doc, '[No se pudieron cargar las fotografías]', size=9, color='94A3B8')
    doc.add_paragraph()


def _sec_indicadores(doc, general_data):
    """Section III – Global acoustic indicators."""
    _heading(doc, 'III.  Indicadores Acústicos Globales')

    laeq = general_data.get('Lp_eqT')
    lexh8 = general_data.get('L_EX_8h')
    et = general_data.get('ET')
    le = general_data.get('LE')
    duration = general_data.get('duration')
    n_audios = general_data.get('cantidad_audios')

    dur_h = f"{duration / 3600:.2f}" if duration else '—'
    dur_s = f"{duration:.0f}" if duration else '—'

    rows = [
        ('LAeq — Nivel Continuo Equivalente', f"{laeq:.2f} dB" if laeq is not None else '—'),
        ('LEX,8h — Exposición normalizada 8 h (NTP ISO 9612)', f"{lexh8:.2f} dB" if lexh8 is not None else '—'),
        ('ET — Exposición Sonora (Pa²·s)', f"{et:.6f}" if et is not None else '—'),
        ('LE — Nivel de Exposición Sonora', f"{le:.2f} dB" if le is not None else '—'),
        ('Duración Total de Medición', f"{dur_s} s ({dur_h} h)"),
        ('N° de Audios Analizados', str(n_audios) if n_audios is not None else '—'),
    ]
    _add_kv_table(doc, rows)

    # Status banner
    limite = LIMITE_SEGURO_DB
    if laeq is not None:
        if laeq >= limite:
            status_txt = f'⚠  LÍMITE EXCEDIDO — LAeq {laeq:.2f} dB ≥ {limite} dB — ACCIÓN OBLIGATORIA'
            status_color = 'FEF2F2'
            text_color = 'B91C1C'
        elif laeq >= 80:
            status_txt = f'▲  NIVEL DE ACCIÓN — LAeq {laeq:.2f} dB — INTERVENCIÓN RECOMENDADA'
            status_color = 'FFFBEB'
            text_color = 'B45309'
        else:
            status_txt = f'✓  AMBIENTE SEGURO — LAeq {laeq:.2f} dB < {limite} dB'
            status_color = 'F0FDF4'
            text_color = '166534'

        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        _set_col_widths(table, _col_widths_twips([1.0]))
        cell = table.rows[0].cells[0]
        _set_cell_bg(cell, status_color)
        _set_cell_borders(cell)
        cell.paragraphs[0].clear()
        run = cell.paragraphs[0].add_run(status_txt)
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(text_color)
        doc.add_paragraph()


def _sec_grafico(doc, chart_img_b64):
    """Section IV – Temporal chart."""
    _heading(doc, 'IV.  Evolución Temporal de Niveles Sonoros')
    img_stream = _safe_b64_to_image(chart_img_b64, max_bytes=MAX_CHART_BYTES)
    if img_stream:
        try:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(img_stream, width=Mm(USABLE_WIDTH_MM))
        except Exception:
            _para(doc, '[Error al insertar el gráfico]', size=9, color='94A3B8')
    else:
        _para(doc, '[Gráfico no disponible — sin imagen proporcionada]', size=9, color='94A3B8')
    doc.add_paragraph()


def _sec_excesos(doc, excesos_data):
    """Section V – Exceedance analysis."""
    _heading(doc, 'V.  Análisis de Excesos del Límite de Referencia')

    if not excesos_data:
        _para(doc, '[Sin datos de excesos disponibles]', size=9, color='94A3B8')
        return

    limite = excesos_data.get('limite_referencia', 85)
    headers = [
        'Evaluación', f'Límite (dB)', 'Pts Medidos', 'Pts Excedieron',
        '% Tiempo', 'Tiempo Exceso', 'Nivel Máx. (dB)', 'N° Audios'
    ]
    eval_seg = excesos_data.get('evaluacion_seguridad', '—')
    pct = excesos_data.get('porcentaje_exceso', 0)
    dur_exc = excesos_data.get('duracion_total_exceso_minutos', 0)
    dur_str = f"{dur_exc} min" if excesos_data.get('puntos_que_excedieron', 0) > 0 else '—'
    max_niv = excesos_data.get('nivel_maximo_registrado', 0)
    max_str = str(max_niv) if excesos_data.get('puntos_que_excedieron', 0) > 0 else '—'

    data_rows = [[
        eval_seg,
        str(limite),
        str(excesos_data.get('total_puntos_medidos', 0)),
        str(excesos_data.get('puntos_que_excedieron', 0)),
        f"{pct}%",
        dur_str,
        max_str,
        str(excesos_data.get('cantidad_audios', 0)),
    ]]
    col_widths = [0.18, 0.09, 0.10, 0.13, 0.10, 0.12, 0.14, 0.10]
    _add_data_table(doc, headers, data_rows, col_widths)

    # Per-audio detail if available
    excesos_por_audio = excesos_data.get('excesos_por_audio', [])
    if excesos_por_audio:
        _para(doc, 'Detalle de Excesos por Audio:', size=9, bold=True)
        h2 = ['Audio ID', 'Pts Excedidos', 'Duración Exceso (s)', 'Nivel Máx. (dB)', 'Estado']
        d2 = [
            [
                str(a.get('audio_id', '—')),
                str(a.get('puntos_excedidos', 0)),
                str(a.get('duracion_exceso_segundos', 0)),
                str(a.get('nivel_maximo', '—')),
                '⚠ Excede' if a.get('puntos_excedidos', 0) > 0 else '✓ OK',
            ]
            for a in excesos_por_audio
        ]
        _add_data_table(doc, h2, d2, [0.18, 0.20, 0.24, 0.20, 0.18])


def _sec_conclusiones(doc, general_data, excesos_data):
    """Section VI – Conclusions and recommendations."""
    _heading(doc, 'VI.  Conclusiones y Recomendaciones')

    laeq = general_data.get('Lp_eqT') if general_data else None
    lexh8 = general_data.get('L_EX_8h') if general_data else None
    pct_exceso = excesos_data.get('porcentaje_exceso', 0) if excesos_data else 0
    max_nivel = excesos_data.get('nivel_maximo_registrado', 0) if excesos_data else 0
    dur_min = excesos_data.get('duracion_total_exceso_minutos', 0) if excesos_data else 0
    limite = LIMITE_SEGURO_DB

    if laeq is None:
        _para(doc, '[No hay datos suficientes para generar conclusiones]', size=9, color='94A3B8')
        return

    # Classification
    if laeq < 70:
        clasificacion = 'AMBIENTE CONTROLADO — SIN RIESGO ACÚSTICO'
        riesgo = 'Los niveles de ruido medidos están muy por debajo del límite de referencia. No se requieren medidas correctivas.'
        status_color = 'F0FDF4'
        text_color = '166534'
    elif laeq < 80:
        clasificacion = 'NIVEL BAJO — PRECAUCIÓN'
        riesgo = 'Los niveles están por debajo del límite pero se recomienda monitoreo periódico como medida preventiva.'
        status_color = 'F0FDF4'
        text_color = '166534'
    elif laeq < limite:
        clasificacion = 'NIVEL DE ACCIÓN — INTERVENCIÓN RECOMENDADA'
        riesgo = f'Los niveles se acercan al límite máximo permisible ({limite} dB). Se recomiendan medidas preventivas.'
        status_color = 'FFFBEB'
        text_color = 'B45309'
    else:
        clasificacion = 'LÍMITE EXCEDIDO — ACCIÓN OBLIGATORIA'
        riesgo = (f'La exposición al ruido supera el límite máximo permisible de {limite} dB '
                  'establecido por la NTP ISO 9612:2010. Se requieren medidas correctivas inmediatas.')
        status_color = 'FEF2F2'
        text_color = 'B91C1C'

    # Banner
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    _set_col_widths(table, _col_widths_twips([1.0]))
    cell = table.rows[0].cells[0]
    _set_cell_bg(cell, status_color)
    _set_cell_borders(cell)
    cell.paragraphs[0].clear()
    run = cell.paragraphs[0].add_run(clasificacion)
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(text_color)
    p2 = cell.add_paragraph(riesgo)
    p2.paragraph_format.space_before = Pt(2)
    for run in p2.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string(text_color)
    doc.add_paragraph()

    # Compliance table
    _para(doc, 'Tabla de Evaluación de Cumplimiento:', size=9, bold=True)
    margin_laeq = laeq - limite
    headers = ['Indicador', 'Valor Medido', 'Límite Ref.', 'Margen', 'Estado']
    data_rows = [
        ['LAeq Global — Nivel continuo equivalente ponderado A',
         f'{laeq:.2f} dB', f'{limite} dB', f'{margin_laeq:+.2f} dB',
         '⚠ EXCEDE' if laeq >= limite else ('▲ PRECAUCIÓN' if laeq >= 80 else '✓ CUMPLE')],
    ]
    if lexh8 is not None:
        margin_lex = lexh8 - limite
        data_rows.append([
            'LEX,8h — Exposición normalizada a 8 h (NTP ISO 9612:2010)',
            f'{lexh8:.2f} dB', f'{limite} dB', f'{margin_lex:+.2f} dB',
            '⚠ EXCEDE' if lexh8 >= limite else ('▲ PRECAUCIÓN' if lexh8 >= 80 else '✓ CUMPLE'),
        ])
    data_rows.append([
        f'Tiempo con Lp,max > {limite} dB',
        f'{pct_exceso}% ({dur_min} min)', '0%', f'+{pct_exceso}%',
        '⚠ ALTO' if pct_exceso > 25 else ('▲ MODERADO' if pct_exceso > 5 else '✓ BAJO'),
    ])
    data_rows.append([
        'Nivel Máximo (Lp,max)',
        f'{max_nivel} dB', '135 dB (Trauma Agudo)', f'{max_nivel - 135:+.1f} dB',
        '⚠ PELIGROSO' if max_nivel > 135 else ('▲ ALTO' if max_nivel > 115 else '✓ OK'),
    ])
    _add_data_table(doc, headers, data_rows, [0.35, 0.18, 0.18, 0.12, 0.17])

    # Recommendations
    _para(doc, 'Recomendaciones Técnicas:', size=10, bold=True)
    if laeq >= limite or (lexh8 is not None and lexh8 >= limite):
        recs = [
            'Protección auditiva individual obligatoria: Implementar tapones o protectores tipo copa con NRR suficiente.',
            'Controles de ingeniería: Evaluar barreras acústicas, aislamiento de fuentes o encerramientos acústicos.',
            'Programa de conservación auditiva: Audiometrías periódicas (mínimo anual) según D.S. N° 005-2012-TR Art. 36.',
            'Re-evaluación urgente: Nueva medición en un plazo no mayor a 30 días post-implementación de controles.',
            'Señalización: Instalar señales de advertencia de ruido conforme a NTP 399.010-1:2004.',
        ]
    elif laeq >= 80:
        recs = [
            'Protección auditiva preventiva: Se recomienda el uso voluntario de protección auditiva.',
            'Monitoreo semestral: Programar evaluaciones acústicas cada 6 meses.',
            'Revisión de equipos: Evaluar mantenimiento preventivo para reducir emisión sonora.',
        ]
    else:
        recs = [
            'Mantener condiciones actuales: El ambiente acústico se encuentra dentro de los límites normativos.',
            'Monitoreo anual: Programar evaluaciones acústicas anuales.',
        ]

    if pct_exceso > 25:
        recs.append(f'Picos frecuentes: El {pct_exceso}% del tiempo registró niveles máximos por encima de {limite} dB.')
    if max_nivel > 115:
        recs.append(f'Protección contra picos: Niveles máximos de {max_nivel} dB. Evaluar protectores auditivos con atenuación impulsiva.')

    for rec in recs:
        p = doc.add_paragraph(style='List Number')
        run = p.add_run(rec)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string('1e293b')

    # Disclaimer
    doc.add_paragraph()
    p_disc = doc.add_paragraph()
    p_disc.paragraph_format.space_before = Pt(4)
    run = p_disc.add_run(
        'Nota: Las conclusiones se derivan de los datos analizados y los criterios de la NTP ISO 9612:2010, '
        'el D.S. N° 005-2012-TR y las directrices de la OIT. Los resultados deben ser validados por un '
        'profesional de salud ocupacional habilitado antes de su uso como sustento legal o médico.'
    )
    run.font.size = Pt(8)
    run.font.italic = True
    run.font.color.rgb = RGBColor.from_string('64748B')
    doc.add_paragraph()


def _sec_datos_tecnicos(doc, audios_data):
    """Section VII – Technical data per audio."""
    _heading(doc, 'VII.  Datos Técnicos por Audio')

    if not audios_data:
        _para(doc, '[Sin datos de audios disponibles]', size=9, color='94A3B8')
        return

    headers = ['Audio', 'Fecha/Hora', 'LAeq (dB)', 'Lp,max (dB)', 'Duración (s)',
               'SR (Hz)', 'ET', 'LE (dB)', 'LI (dB)', 'LJ (dB)', 'Estado']
    col_widths = [0.06, 0.14, 0.08, 0.09, 0.09, 0.07, 0.09, 0.08, 0.08, 0.08, 0.10]

    data_rows = []
    for audio in audios_data:
        gr = audio.get('global_result') or {}
        # Check exceedance
        excede = False
        for pt in (audio.get('detailed_results') or []):
            if float(pt.get('Lp_max', 0)) > LIMITE_SEGURO_DB:
                excede = True
                break
        et_val = gr.get('ET')
        et_str = f"{et_val:.2e}" if et_val is not None else '—'
        data_rows.append([
            str(audio.get('audio_id', '—')),
            _fmt_ts(audio.get('timestamp', '')),
            _fmt_num(gr.get('Lp_eqT')),
            _fmt_num(gr.get('Lp_max')),
            _fmt_num(gr.get('duration')),
            _fmt_num(gr.get('sample_rate'), decimals=0),
            et_str,
            _fmt_num(gr.get('LE')),
            _fmt_num(gr.get('LI')),
            _fmt_num(gr.get('LJ')),
            '⚠ Excede' if excede else '✓ OK',
        ])

    _add_data_table(doc, headers, data_rows, col_widths)


def _sec_metodologia(doc):
    """Section VIII – Methodology (abbreviated)."""
    _heading(doc, 'VIII.  Metodología y Proceso Técnico')

    paras = [
        ('Marco Normativo',
         'Este análisis sigue los criterios de la NTP ISO 9612:2010 (Acústica — '
         'Determinación de la exposición al ruido en el trabajo), el D.S. N° 005-2012-TR '
         '(Reglamento de SST del Perú) y las directrices de la OIT sobre sistemas de gestión '
         'de la seguridad y salud en el trabajo.'),
        ('Indicadores Calculados',
         'LAeq (Nivel de presión sonora continuo equivalente ponderado A), LEX,8h '
         '(Exposición normalizada a 8 horas), Lp,max (Nivel de pico de presión sonora), '
         'ET (Exposición sonora acumulada) y LE (Nivel de exposición sonora).'),
        ('Proceso de Análisis',
         'Los archivos de audio WAV son procesados segundo a segundo. Se calcula el '
         'promedio energético (integración temporal) de la presión sonora ponderada A '
         'para obtener el LAeq global del período completo de medición.'),
        ('Límite de Referencia',
         'El límite máximo permisible (LMP) utilizado es de 85 dB para una jornada laboral '
         'de 8 horas, conforme a la NTP ISO 9612:2010 y legislación peruana vigente.'),
    ]
    for title, text in paras:
        _para(doc, title, size=10, bold=True, space_after=2)
        _para(doc, text, size=9, color='374151', space_after=8)


def _sec_glosario(doc):
    """Section IX – Glossary."""
    _heading(doc, 'IX.  Glosario de Términos')
    terms = [
        ('LAeq', 'Nivel de presión sonora continuo equivalente ponderado A. Representa el nivel de ruido constante que contiene la misma energía sonora que el ruido fluctuante medido.'),
        ('LEX,8h', 'Nivel de exposición al ruido normalizado para una jornada laboral de 8 horas (NTP ISO 9612:2010).'),
        ('Lp,max', 'Nivel de pico de presión sonora — valor máximo instantáneo registrado durante la medición.'),
        ('LMP', 'Límite Máximo Permisible — valor de LAeq por encima del cual la exposición sin protección supone riesgo de daño auditivo.'),
        ('ET', 'Exposición sonora acumulada expresada en Pa²·s.'),
        ('LE', 'Nivel de exposición sonora expresado en dB, equivalente logarítmico de ET.'),
        ('NTP ISO 9612:2010', 'Norma Técnica Peruana que establece el método para determinar la exposición al ruido ocupacional.'),
    ]
    rows = [(t, d) for t, d in terms]
    _add_kv_table(doc, rows, col_widths=[0.22, 0.78])


def _sec_referencias(doc):
    """Section X – References."""
    _heading(doc, 'X.  Referencias Bibliográficas')
    refs = [
        'NTP ISO 9612:2010 — Acústica. Determinación de la exposición al ruido en el lugar de trabajo. INACAL, Lima.',
        'D.S. N° 005-2012-TR — Reglamento de la Ley de Seguridad y Salud en el Trabajo. El Peruano, Lima.',
        'OIT (2009) — Directrices sobre sistemas de gestión de la seguridad y la salud en el trabajo (ILO-OSH 2001). OIT, Ginebra.',
        'ISO 1999:2013 — Acoustics — Estimation of noise-induced hearing loss. ISO, Ginebra.',
        'NIOSH (1998) — Criteria for a Recommended Standard: Occupational Noise Exposure (Revised Criteria 1998). CDC/NIOSH, Cincinnati.',
    ]
    for i, ref in enumerate(refs, 1):
        p = doc.add_paragraph(style='List Number')
        run = p.add_run(ref)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string('374151')


# ─── Number / date formatters ─────────────────────────────────────────────────

def _fmt_num(val, decimals=2):
    if val is None:
        return '—'
    try:
        if decimals == 0:
            return str(int(float(val)))
        return f'{float(val):.{decimals}f}'
    except Exception:
        return str(val)


def _fmt_ts(ts):
    if not ts:
        return '—'
    try:
        # Try to parse ISO 8601 and format nicely
        from dateutil import parser as dp
        dt = dp.parse(str(ts))
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return str(ts)[:19]  # fallback: trim to datetime portion


# ─── Public API ───────────────────────────────────────────────────────────────

def build_informe_docx(
    esc_data,
    mic_data,
    general_data,
    excesos_data,
    audios_data,
    chart_img_b64=None,
    fotos_paths=None,
    secciones=None,
    fecha_generacion=None,
):
    """
    Build a complete informe DOCX and return a BytesIO buffer.

    Parameters
    ----------
    esc_data : dict        Scenario metadata (from /escenarios/<id>/detalle)
    mic_data : dict|None   Microphone metadata (may be None)
    general_data : dict    General acoustic analysis
    excesos_data : dict    Exceedance analysis
    audios_data : list     Per-audio technical data
    chart_img_b64 : str|None  Base-64 encoded PNG of the Highcharts chart
    fotos_paths : list|None   List of (BytesIO | None) streams, one per photo
    secciones : dict|None  Which sections to include (all True by default)
    fecha_generacion : str  Human-readable generation timestamp

    Returns
    -------
    BytesIO with the DOCX content (position at 0).
    """
    if secciones is None:
        secciones = {}
    _inc = lambda k: secciones.get(k, True)

    if fecha_generacion is None:
        fecha_generacion = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%d %H:%M UTC')

    if fotos_paths is None:
        fotos_paths = []

    doc = _new_document()

    # Cover page
    _sec_portada(doc, esc_data, mic_data, fecha_generacion)

    # I – Scenario information (always)
    _sec_escenario(doc, esc_data, mic_data)

    # II – Photos
    if _inc('fotos'):
        _sec_fotos(doc, fotos_paths)

    # III – Global indicators (always)
    _sec_indicadores(doc, general_data)

    # IV – Chart
    if _inc('grafico'):
        _sec_grafico(doc, chart_img_b64)

    # V – Exceedance analysis (always)
    _sec_excesos(doc, excesos_data)

    # VI – Conclusions (always)
    _sec_conclusiones(doc, general_data, excesos_data)

    # VII – Technical data
    if _inc('datos'):
        _sec_datos_tecnicos(doc, audios_data)

    # VIII – Methodology
    if _inc('metodologia'):
        _sec_metodologia(doc)

    # IX – Glossary
    if _inc('glosario'):
        _sec_glosario(doc)

    # X – References
    if _inc('referencias'):
        _sec_referencias(doc)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_informe_combinado_docx(informes, fecha_generacion=None):
    """
    Build a combined DOCX with multiple scenario reports.

    Parameters
    ----------
    informes : list of dicts, each with keys:
        esc_data, mic_data, general_data, excesos_data, audios_data,
        chart_img_b64, fotos_paths, secciones
    fecha_generacion : str

    Returns
    -------
    BytesIO with the combined DOCX content (position at 0).
    """
    if fecha_generacion is None:
        fecha_generacion = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%d %H:%M UTC')

    if not informes:
        doc = _new_document()
        _para(doc, 'Sin informes para combinar.', size=11)
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    doc = _new_document()

    # Combined cover page
    doc.add_paragraph()
    doc.add_paragraph()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run('INFORME TÉCNICO COMBINADO')
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string('64748B')

    p_main = doc.add_paragraph()
    p_main.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_main.add_run('Análisis de Exposición al Ruido Ocupacional')
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string('0F172A')

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_sub.add_run(f'{len(informes)} Escenario(s) incluido(s)')
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor.from_string('475569')

    doc.add_paragraph()
    p_gen = doc.add_paragraph()
    p_gen.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_gen.add_run(f'Generado: {fecha_generacion}')
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string('94A3B8')

    # Index
    doc.add_paragraph()
    _heading(doc, 'Escenarios incluidos:', level=2)
    for i, inf in enumerate(informes, 1):
        esc = inf.get('esc_data', {})
        nombre = esc.get('nombre') or f"Escenario #{esc.get('id', i)}"
        ubicacion = esc.get('ubicacion', '')
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. {nombre}')
        run.font.size = Pt(10)
        run.font.bold = True
        if ubicacion:
            p2 = doc.add_paragraph()
            run2 = p2.add_run(f'   Ubicación: {ubicacion}')
            run2.font.size = Pt(9)
            run2.font.color.rgb = RGBColor.from_string('64748B')

    _add_page_break(doc)

    # Individual scenario sections
    for i, inf in enumerate(informes):
        secciones = inf.get('secciones') or {}
        _inc = lambda k: secciones.get(k, True)

        esc_data = inf.get('esc_data', {})
        mic_data = inf.get('mic_data')
        general_data = inf.get('general_data', {})
        excesos_data = inf.get('excesos_data', {})
        audios_data = inf.get('audios_data', [])
        chart_img_b64 = inf.get('chart_img_b64')
        fotos_paths = inf.get('fotos_paths') or []

        # Scenario separator heading
        nombre = esc_data.get('nombre') or f"Escenario #{esc_data.get('id', i + 1)}"
        p_sep = doc.add_paragraph()
        p_sep.paragraph_format.space_before = Pt(0)
        run = p_sep.add_run(f'─── Escenario {i + 1}: {nombre} ───')
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string('1e293b')
        doc.add_paragraph()

        _sec_escenario(doc, esc_data, mic_data)
        if _inc('fotos'):
            _sec_fotos(doc, fotos_paths)
        _sec_indicadores(doc, general_data)
        if _inc('grafico'):
            _sec_grafico(doc, chart_img_b64)
        _sec_excesos(doc, excesos_data)
        _sec_conclusiones(doc, general_data, excesos_data)
        if _inc('datos'):
            _sec_datos_tecnicos(doc, audios_data)

        if i < len(informes) - 1:
            _add_page_break(doc)

    # Shared methodology/glossary/references at the end
    _add_page_break(doc)
    _sec_metodologia(doc)
    _sec_glosario(doc)
    _sec_referencias(doc)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
