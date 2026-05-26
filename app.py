"""
app.py  ·  Facts
---------------------
Interfaz Streamlit para el extractor de facturas.
Ejecutar desde la raíz del proyecto:
    streamlit run app.py
"""

import sys
import tempfile
from pathlib import Path

import streamlit as st

# ── Configuración de página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Facts",
    page_icon="assets/favicon.png" if Path("assets/favicon.png").exists() else None,
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Estilos ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

/* Reset y base */
html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Fondo */
.stApp {
    background-color: #F7F7F5;
}

/* Ocultar elementos de Streamlit que no necesitamos */
#MainMenu, footer, header { visibility: hidden; }
.block-container {
    padding-top: 3rem;
    padding-bottom: 3rem;
    max-width: 760px;
}

/* Cabecera */
.factura-header {
    margin-bottom: 2.5rem;
    padding-bottom: 1.5rem;
    border-bottom: 1px solid #E0DED8;
}
.factura-header h1 {
    font-size: 1.75rem;
    font-weight: 600;
    color: #1A1A18;
    letter-spacing: -0.02em;
    margin: 0 0 0.35rem 0;
}
.factura-header p {
    font-size: 0.925rem;
    color: #6B6B65;
    margin: 0;
    font-weight: 400;
}

/* Zona de carga */
.upload-zone {
    background: #FFFFFF;
    border: 1.5px dashed #D0CEC8;
    border-radius: 8px;
    padding: 2rem;
    margin-bottom: 1.25rem;
    transition: border-color 0.2s;
}
.upload-zone:hover {
    border-color: #1A1A18;
}

/* Etiqueta de archivos listos */
.files-ready {
    background: #F0F5F0;
    border: 1px solid #C8DCC8;
    border-radius: 6px;
    padding: 0.65rem 1rem;
    font-size: 0.875rem;
    color: #2D5A2D;
    font-weight: 500;
    margin-bottom: 1.25rem;
}

/* Botón principal */
.stButton > button {
    background-color: #1A1A18 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.01em !important;
    padding: 0.6rem 1.5rem !important;
    height: auto !important;
    transition: background-color 0.15s !important;
}
.stButton > button:hover {
    background-color: #333330 !important;
}

/* Botón de descarga */
.stDownloadButton > button {
    background-color: #FFFFFF !important;
    color: #1A1A18 !important;
    border: 1.5px solid #1A1A18 !important;
    border-radius: 6px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    width: 100% !important;
    padding: 0.65rem !important;
    margin-top: 0.5rem !important;
}
.stDownloadButton > button:hover {
    background-color: #1A1A18 !important;
    color: #FFFFFF !important;
}

/* Métricas */
[data-testid="metric-container"] {
    background: #FFFFFF;
    border: 1px solid #E8E6E0;
    border-radius: 8px;
    padding: 1rem 1.25rem !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.8rem !important;
    color: #6B6B65 !important;
    font-weight: 500 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.05em !important;
}
[data-testid="stMetricValue"] {
    font-size: 2rem !important;
    font-weight: 600 !important;
    color: #1A1A18 !important;
    font-family: 'DM Mono', monospace !important;
}

/* Spinner */
.stSpinner > div {
    border-top-color: #1A1A18 !important;
}

/* Separador */
hr {
    border: none;
    border-top: 1px solid #E0DED8;
    margin: 1.75rem 0;
}

/* Alertas */
.stAlert {
    border-radius: 6px !important;
    font-size: 0.875rem !important;
}

/* Log de errores */
.error-log {
    background: #FDF5F5;
    border: 1px solid #E8C8C8;
    border-radius: 6px;
    padding: 1rem 1.25rem;
    font-family: 'DM Mono', monospace;
    font-size: 0.8rem;
    color: #8B2E2E;
    margin-top: 1rem;
}
.error-log-title {
    font-family: 'DM Sans', sans-serif;
    font-weight: 600;
    font-size: 0.8rem;
    color: #8B2E2E;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 0.5rem;
}

/* Sección de resultados */
.results-header {
    font-size: 0.8rem;
    font-weight: 600;
    color: #6B6B65;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ── Import del motor ──────────────────────────────────────────────────────────
try:
    from core.procesador_lote import ProcesadorLote
    _motor_ok = True
except ModuleNotFoundError as e:
    _motor_ok = False
    _motor_error = str(e)

# ── Cabecera ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="factura-header">
    <h1>FactsAI</h1>
    <p>Extrae datos fiscales de tus facturas y exporta a Excel listo para contabilidad.</p>
</div>
""", unsafe_allow_html=True)

# ── Error de importación ──────────────────────────────────────────────────────
if not _motor_ok:
    st.error(
        f"No se pudo cargar el motor de extracción: `{_motor_error}`  \n"
        "Asegúrate de ejecutar Streamlit desde la raíz del proyecto:  \n"
        "`streamlit run app.py`"
    )
    st.stop()

# ── Zona de carga ─────────────────────────────────────────────────────────────
archivos_subidos = st.file_uploader(
    "Facturas",
    type=["pdf", "jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="Formatos aceptados: PDF, JPG, PNG, WEBP",
)

if archivos_subidos:
    n = len(archivos_subidos)
    plural = "archivos" if n > 1 else "archivo"
    st.markdown(
        f'<div class="files-ready">{n} {plural} listo{"s" if n > 1 else ""} para procesar</div>',
        unsafe_allow_html=True,
    )

    if st.button("Procesar facturas", type="primary", use_container_width=False):

        with st.spinner("Analizando documentos..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Guardar archivos subidos en disco temporal
                rutas = []
                for archivo in archivos_subidos:
                    ruta = Path(tmpdir) / archivo.name
                    ruta.write_bytes(archivo.getbuffer())
                    rutas.append(ruta)

                # Ejecutar el motor
                procesador = ProcesadorLote()
                informe    = procesador.procesar_archivos(rutas)

                # Generar Excel
                ruta_excel = Path(tmpdir) / "facturas.xlsx"
                informe._construir_excel(ruta_excel)
                excel_bytes = ruta_excel.read_bytes()

        # ── Métricas ──────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown('<div class="results-header">Resumen del lote</div>', unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        col1.metric("Procesadas",  informe.total)
        col2.metric("Válidas",     informe.validas)
        col3.metric("Con errores", informe.con_error)

        # ── Descarga ──────────────────────────────────────────────────────────
        st.markdown("---")
        st.download_button(
            label="Descargar Excel",
            data=excel_bytes,
            file_name=f"facturas_{informe.timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # ── Log de errores ────────────────────────────────────────────────────
        errores = [f for f in informe.filas if f.estado == "error"]
        if errores:
            items = "".join(
                f"<div>{e.archivo}: {e.error_log}</div>" for e in errores
            )
            st.markdown(
                f'<div class="error-log">'
                f'<div class="error-log-title">Archivos con error</div>'
                f'{items}'
                f'</div>',
                unsafe_allow_html=True,
            )
