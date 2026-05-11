import streamlit as st
import fitz  # PyMuPDF
fitz.TOOLS.mupdf_display_errors(False)
import google.generativeai as genai
import io
import json
import re
import time
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 🎨 CONFIGURACIÓN DE PÁGINA Y ESTILOS (VERDE MÉDICO HUJ) ---
st.set_page_config(page_title="BioCardio Clinical Assistant", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    .stApp {
        background-color: #f4f7f4;
    }
    /* Estilo para las etiquetas de carga de archivos */
    .stFileUploader label {
        font-weight: bold !important;
        color: #1a1a1a !important;
        font-size: 19px !important;
        display: block;
        margin-bottom: 12px !important;
    }
    .step-container {
        background-color: white;
        padding: 25px 30px;
        border-radius: 15px;
        border-left: 8px solid #007d32;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        margin-bottom: 35px;
    }
    .step-header {
        color: #007d32;
        font-size: 24px;
        font-weight: bold;
        margin-top: 15px;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .step-explanation {
        color: #1a1a1a;
        font-size: 16px;
        font-weight: 500;
        margin-bottom: 20px;
        font-style: italic;
        border-bottom: 2px solid #eee;
        padding-bottom: 10px;
    }
    .stButton>button {
        background-color: #007d32 !important;
        color: white !important;
        border-radius: 12px !important;
        padding: 15px 30px !important;
        font-size: 18px !important;
        font-weight: bold !important;
        width: 100%;
        transition: 0.3s;
        border: none !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 📄 GENERADOR DEL WORD ---
def crear_documento_word_pro(datos_json):
    try:
        match = re.search(r'\{.*\}', datos_json, re.DOTALL)
        datos = json.loads(match.group(0)) if match else json.loads(datos_json)
    except Exception as e:
        datos = {"visita": "Error en formato", "procedimientos_finales": []}

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    
    doc.add_heading(f"HOJA DE VISITA: {datos.get('visita', 'N/A')}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER

    info_table = doc.add_table(rows=2, cols=2)
    info_table.cell(0, 0).text = "Protocolo / Centro:\n_________________________"
    info_table.cell(0, 1).text = "ID del Paciente / Iniciales:\n_________________________"
    
    doc.add_paragraph("\n⚠️ NOTA: Si el paciente no acude a la visita, rellenar igualmente el eCRF.").runs[0].bold = True

    procedimientos = datos.get("procedimientos_finales", [])
    categorias_ordenadas = [
        ("PRE-DOSIS / PRE-INFUSIÓN", "pre_dosis"),
        ("EXTRACCIONES DE LABORATORIO CENTRAL", "laboratorio"),
        ("ADMINISTRACIÓN DE TRATAMIENTO / INFUSIÓN", "administracion"),
        ("POST-DOSIS / SEGUIMIENTO", "post_dosis"),
        ("PROCEDIMIENTOS CONTINUOS", "continuos"),
        ("OTROS PROCEDIMIENTOS", "general")
    ]

    for titulo_cat, id_cat in categorias_ordenadas:
        items = [p for p in procedimientos if p.get("categoria") == id_cat]
        if items:
            doc.add_heading(titulo_cat, level=1)
            for item in items:
                p = doc.add_paragraph()
                p.add_run(f"[  ] {item.get('procedimiento')} ").bold = True
                
                proc_nombre = item.get('procedimiento', '').lower()

                if any(x in proc_nombre for x in ["ecg", "electrocardiograma", "12-lead"]):
                    doc.add_paragraph("• Posición supino. FC: ... PR: ... QRS: ... QT: ... QTc: ...", style='List Bullet')
                
                if any(x in proc_nombre for x in ["6 min", "6-min", "6mwt", "walk", "marcha"]):
                    test_p = doc.add_paragraph()
                    test_p.paragraph_format.left_indent = Pt(20)
                    test_p.add_run("• Realizar tras muestras/cuestionarios y antes de infusión.\n• Course: 20m / Name: 6MWT 2F.\n• Rellenar todas las casillas (N/D si no aplica).").italic = True

                if item.get("detalles"):
                    det_p = doc.add_paragraph(item.get("detalles"))
                    det_p.paragraph_format.left_indent = Pt(20)

                if "signos vitales" in proc_nombre or "vital signs" in proc_nombre:
                    tiempos = item.get("tiempos_especificos", ["________________", "________________"])
                    table = doc.add_table(rows=len(tiempos)+1, cols=8)
                    table.style = 'Table Grid'
                    headers = ["Tiempo", "Hora", "PA Sist.", "PA Diast.", "Pulso", "Resp", "O2 (%)", "Temp."]
                    for i, head in enumerate(headers): table.cell(0, i).text = head
                    for i, t_val in enumerate(tiempos): table.cell(i+1, 0).text = t_val

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def extraer_texto_paginas(pdf_bytes, paginas_str=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    try:
        p_set = set()
        for b in paginas_str.split(','):
            if '-' in b:
                s, e = b.split('-')
                p_set.update(range(int(s), int(e)+1))
            else: p_set.add(int(b))
        for p in sorted(list(p_set)):
            if 0 < p <= len(doc):
                texto += f"\n--- PÁGINA {p} ---\n" + doc[p-1].get_text("text") + "\n"
    except: texto = "Error en lectura de páginas."
    return texto

# --- 🖥️ INTERFAZ WEB ---
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try: st.image("huj.png", width=140)
    except: st.markdown("## 🏥")
with col_titulo:
    st.markdown("<h1 style='color: #007d32; margin-top: 10px;'>BioCardio Clinical Assistant</h1>", unsafe_allow_html=True)
    st.write("Hospital Universitario de Jaén | Gestión de Precisión v3.7")

# --- BARRA LATERAL (RECUPERADA) ---
with st.sidebar:
    st.header("🔑 Configuración")
    st.markdown("""
    **¿Cómo obtener tu clave?**
    1. Entra en [Google AI Studio](https://aistudio.google.com/app/apikey).
    2. Pulsa en **"Create API Key"**.
    3. Copia el código y pégalo aquí abajo:
    """)
    api_key = st.text_input("Introduce tu API Key personal:", type="password")
    
    st.divider()
    modelo_seleccionado = st.selectbox("Modelo de IA:", ["gemini-3-flash"], help="Modelo optimizado para máxima precisión y velocidad.")
    st.info("Nota: Los documentos subidos se borran automáticamente al cerrar la sesión.")

# PASOS
st.markdown('<div class="step-header">📄 Paso 1: Documentación</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">Sube el Protocolo y los Manuales de Laboratorio:</div>', unsafe_allow_html=True)
    f_proto = st.file_uploader("1. SUBIR PROTOCOLO (Apartado SoE)", type=["pdf"])
    f_labs = st.file_uploader("2. SUBIR MANUALES DE LABORATORIO", type=["pdf"], accept_multiple_files=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="step-header">🔍 Paso 2: Configuración de la Visita</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">Define las páginas y la visita exacta:</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    p_tabla = c1.text_input("Páginas SoE (Tabla):", "11-16")
    v_proto = c1.text_input("Nombre Visita en Tabla:", "Visit 17")
    p_assessments = c2.text_input("Páginas Study Assessments:", "40-60")
    v_lab = c2.text_input("Nombre Visita en Lab:", "Visit 17")
    st.markdown('</div>', unsafe_allow_html=True)

if st.button("✨ GENERAR HOJA DE VISITA SIN OMISIONES"):
    if not api_key or not f_proto:
        st.error("⚠️ Debes introducir la API Key y subir el Protocolo.")
    else:
        with st.spinner("Escaneando tabla minuciosamente..."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(modelo_seleccionado)
                p_bytes = f_proto.read()
                t_soe = extraer_texto_paginas(p_bytes, p_tabla)
                t_ass = extraer_texto_paginas(p_bytes, p_assessments)
                
                prompt = f"""
                Misión: Traslada TODA la información de la columna '{v_proto}' sin omitir ni un solo punto.
                
                TABLA SOE: {t_soe}
                DETALLES TÉCNICOS: {t_ass}

                REGLAS DE ORO:
                1. Escanea la columna '{v_proto}' de arriba a abajo. Cualquier fila con una marca (X, punto, asterisco) DEBE incluirse.
                2. ES OBLIGATORIO incluir (si están marcados): 6-Minute Walk Test, NYHA Class, Physical Exam, Medical Review, Echocardiogram.
                3. Si el término '6-minute' o 'walk' aparece en las páginas de la tabla, inclúyelo obligatoriamente en 'pre_dosis'.
                4. No resumas nombres. Extrae detalles de 'Study Assessments'.
                
                JSON: {{ "visita": "{v_proto}", "procedimientos_finales": [...] }}
                """
                res = model.generate_content(prompt)
                st.download_button("⬇️ DESCARGAR WORD", crear_documento_word_pro(res.text), f"Checklist_{v_proto}.docx")
                st.success("✅ ¡Generado! Ya puedes descargar.")
            except Exception as e: st.error(f"Error: {e}")
