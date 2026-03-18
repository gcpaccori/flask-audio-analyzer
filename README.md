# Flask Audio Analyzer

Sistema de análisis de exposición al ruido ocupacional basado en Flask.

## Instalación

```bash
pip install -r requirements.txt
python app.py
```

## Exportación a Word (.docx)

Desde la pantalla de **informe finalizado** (análisis de escenario) se pueden descargar dos tipos de informes Word editables:

### Botones disponibles en el modal "Imprimir / Exportar"

| Botón | Acción |
|---|---|
| **Exportar Word (.docx)** | Descarga el informe del escenario actual como `.docx` editable |
| **Exportar Word combinado (.docx)** | Descarga un `.docx` único que incluye todos los escenarios seleccionados |

### Formato del documento DOCX

- **Página**: A4 (210 × 297 mm)
- **Márgenes**: izquierda 15 mm, derecha 15 mm, superior 15 mm, inferior 20 mm
- **Estructura de secciones**:
  1. Portada (título, nombre del escenario, fechas, micrófono)
  2. I. Información del Escenario
  3. II. Fotografías del Lugar de Medición
  4. III. Indicadores Acústicos Globales
  5. IV. Evolución Temporal de Niveles Sonoros (gráfico PNG)
  6. V. Análisis de Excesos del Límite de Referencia
  7. VI. Conclusiones y Recomendaciones
  8. VII. Datos Técnicos por Audio
  9. VIII. Metodología y Proceso Técnico
  10. IX. Glosario de Términos
  11. X. Referencias Bibliográficas

### Nombres de archivos generados

- Individual: `Informe_<NombreEscenario>_<YYYYMMDD>.docx`
- Combinado: `Informe_Comb_<YYYYMMDD>.docx`

### Endpoints REST

| Método | Ruta | Descripción |
|---|---|---|
| `POST` | `/escenarios/<id>/microfono/<mic_id>/exportar-docx` | Informe individual |
| `POST` | `/escenarios/exportar-docx-combinado` | Informe combinado |

#### Payload individual

```json
{
  "chart_image": "<base64 PNG del gráfico Highcharts, sin prefijo data:>",
  "secciones": {
    "fotos": true,
    "grafico": true,
    "datos": true,
    "metodologia": true,
    "glosario": true,
    "referencias": true
  }
}
```

#### Payload combinado

```json
{
  "escenarios": [
    {
      "escenario_id": 1,
      "microfono_id": 2,
      "chart_image": "<base64 PNG>",
      "secciones": {}
    }
  ]
}
```

### Dependencias adicionales

- `python-docx>=1.1.2` — generación de documentos Word editables

## Normas de referencia

- **NTP ISO 9612:2010** — Acústica. Determinación de la exposición al ruido en el lugar de trabajo.
- **D.S. N° 005-2012-TR** — Reglamento de la Ley de Seguridad y Salud en el Trabajo (Perú).
- **OIT ILO-OSH 2001** — Directrices sobre sistemas de gestión de la SST.
