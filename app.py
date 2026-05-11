import streamlit as st
import fitz  # PyMuPDF
fitz.TOOLS.mupdf_display_errors(False)
import google.generativeai as genai
import io
import json
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 🎨 CONFIGURACIÓN Y ESTILOS MEJORADOS ---
st.set_page_config(page_title="BioCardio Clinical Assistant", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f4; }
    
    /* Etiquetas de carga más visibles */
    .stFileUploader label {
        font-weight: bold !important;
        color: #1a1a1a !important;
        font-size: 18px !important;
        margin-bottom: 10px !important;
    }
    
    .step-header {
        color: #007d32;
        font-size: 26px;
        font-weight: bold;
        margin-top: 25px;
        margin-bottom: 10px;
    }
    
    .step-container {
        background-color: white;
        padding: 30px;
        border-radius: 15px;
        border-left: 10px solid #007d32;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin-bottom: 40px;
    }
    
    .step-explanation {
        color: #333;
        font-size: 17px;
        font-weight: 500;
        margin-bottom: 25px;
        padding-bottom: 12px;
        border-bottom: 2px solid #f0f0f0;
    }

    .stButton>button {
        background-color: #007d32 !important;
        height: 60px !important;
        font-size: 20px !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- (Funciones de Word y Extracción se mantienen igual para asegurar estabilidad) ---
def crear_documento_word_pro(datos_json):
    try:
        match = re.search(r'\{.*\}', datos_json, re.DOTALL)
        datos = json.loads(match.group(0)) if match else json.loads(datos_json)
    except:
        datos = {"visita": "Error", "procedimientos_finales": []}

    doc = Document()
    doc.add_heading(f"HOJA DE VISITA: {datos.get('visita', 'N/A')}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    procedimientos = datos.get("procedimientos_finales", [])
    categorias = [
        ("PRE-DOSIS / PRE-INFUSIÓN", "pre_dosis"),
        ("EXTRACCIONES DE LABORATORIO", "laboratorio"),
        ("ADMINISTRACIÓN", "administracion"),
        ("POST-DOSIS", "post_dosis"),
        ("OTROS", "general")
    ]

    for titulo, cat_id in categorias:
        items = [p for p in procedimientos if p.get("categoria") == cat_id]
        if items:
            doc.add_heading(titulo, level=1)
            for item in items:
                p = doc.add_paragraph()
                p.add_run(f"[  ] {item.get('procedimiento')}").bold = True
                
                nombre = item.get('procedimiento', '').lower()
                # Lógica especial para ECG y 6MWT en el Word
                if any(x in nombre for x in ["ecg", "electrocardiograma", "12-lead"]):
                    doc.add_paragraph("• Posición supino\n• FC: ... PR: ... QRS: ... QT: ... QTc: ...", style='List Bullet')
                if any(x in nombre for x in ["6 min", "6mwt", "marcha"]):
                    doc.add_paragraph("• Realizar antes de infusión.\n• Course Length: 20m. Course Name: 6MWT 2F.\n• Rellenar todas las casillas.", style='List Bullet')
                
                if item.get("detalles"):
                    doc.add_paragraph(item.get("detalles"), style='List Bullet')
    
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
                texto += doc[p-1].get_text("text") + "\n"
    except: texto = "Error en lectura de páginas."
    return texto

# --- 🖥️ INTERFAZ ---
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try: st.image("huj.png", width=150)
    except: st.title("🏥")
with col_titulo:
    st.markdown("<h1 style='color: #007d32; margin: 0;'>BioCardio Clinical Assistant</h1>", unsafe_allow_html=True)
    st.write("Hospital Universitario de Jaén - v3.5 (Alta Precisión)")

with st.sidebar:
    st.header("🔑 Configuración")
    api_key = st.text_input("API Key:", type="password")
    modelo = st.selectbox("Cerebro de la IA:", ["gemini-1.5-pro"], help="Usa 'Pro' para visitas difíciles como la V17.")

# PASO 1
st.markdown('<div class="step-header">📄 Paso 1: Documentación</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">Sube los archivos PDF oficiales aquí:</div>', unsafe_allow_html=True)
    f_proto = st.file_uploader("1. SUBIR PROTOCOLO (TABLA SoE)", type=["pdf"])
    f_labs = st.file_uploader("2. SUBIR MANUALES DE LABORATORIO", type=["pdf"], accept_multiple_files=True)
    st.markdown('</div>', unsafe_allow_html=True)

# PASO 2
st.markdown('<div class="step-header">🔍 Paso 2: Configuración de la Visita</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">Indica los datos de la visita para que la IA los localice:</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    p_tabla = c1.text_input("Páginas de la Tabla SoE (ej: 11-15):", "11-15")
    v_proto = c1.text_input("Nombre de la Visita en Tabla (ej: Visit 17):", "Visit 17")
    p_ass = c2.text_input("Páginas de Study Assessments (ej: 40-60):", "40-60")
    v_lab = c2.text_input("Nombre de la Visita en Lab (ej: Visit 17):", "Visit 17")
    st.markdown('</div>', unsafe_allow_html=True)

# PASO 3
if st.button("✨ GENERAR HOJA DE VISITA COMPLETA"):
    if not api_key or not f_proto:
        st.error("Falta la API Key o el Protocolo.")
    else:
        with st.spinner("Analizando tabla minuciosamente..."):
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(modelo)
            
            t_soe = extraer_texto_paginas(f_proto.read(), p_tabla)
            t_ass = extraer_texto_paginas(f_proto.read(), p_ass)
            
            prompt = f"""
            Analiza esta tabla de protocolo clínico para la visita '{v_proto}'.
            
            TABLA SOE:
            {t_soe}
            
            ASSESSMENTS:
            {t_ass}
            
            INSTRUCCIONES DE OBLIGADO CUMPLIMIENTO:
            1. Busca la columna '{v_proto}'. Si hay marcas (X, puntos, asteriscos), extrae el procedimiento.
            2. NO PUEDES OMITIR: '6-minute walk test', '6MWT', 'NYHA class', 'Physical Examination', '12-lead ECG', 'Echocardiogram'.
            3. Si el texto menciona '6-minute walk' en cualquier parte de la tabla asignada, INCLÚYELO sí o sí.
            4. Clasifica todo en: pre_dosis, laboratorio, administracion, post_dosis.
            
            JSON FORMAT:
            {{ "visita": "{v_proto}", "procedimientos_finales": [ {{"procedimiento": "...", "categoria": "...", "detalles": "..."}} ] }}
            """
            
            res = model.generate_content(prompt)
            st.download_button("⬇️ DESCARGAR WORD", crear_documento_word_pro(res.text), f"Visita_{v_proto}.docx")
