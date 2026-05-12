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

# --- 🎨 CONFIGURACIÓN DE PÁGINA Y ESTILOS ---
st.set_page_config(page_title="BioCardio Clinical Assistant", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f4; }
    .stFileUploader label { font-weight: bold !important; color: #1a1a1a !important; font-size: 19px !important; display: block; margin-bottom: 12px !important; }
    .step-container { background-color: white; padding: 25px 30px; border-radius: 15px; border-left: 8px solid #007d32; box-shadow: 0 4px 12px rgba(0,0,0,0.08); margin-bottom: 35px; }
    .step-header { color: #007d32; font-size: 24px; font-weight: bold; margin-top: 15px; margin-bottom: 10px; }
    .step-explanation { color: #555; font-size: 16px; margin-bottom: 20px; font-style: italic; border-bottom: 1px solid #e0e0e0; padding-bottom: 10px; }
    .stButton>button { background-color: #007d32 !important; color: white !important; border-radius: 12px !important; padding: 15px 30px !important; font-size: 18px !important; font-weight: bold !important; width: 100%; transition: 0.3s; border: none !important; }
    .stButton>button:hover { background-color: #005a24 !important; box-shadow: 0 4px 15px rgba(0,125,50,0.3); }
    </style>
""", unsafe_allow_html=True)

# --- 📋 LISTA MAESTRA DE PROCEDIMIENTOS ---
PROCEDIMIENTOS_MAESTROS = """
1. Informed consent / Eligibility assessment
2. Demographics / Medical history
3. Full physical examination
4. Symptom-directed physical examination
5. Height and body weight
6. Karnofsky Performance Status score
7. NYHA Functional Classification
8. 6-Minute Walk Test (6MWT)
9. Vital signs
10. 12-lead ECG (Single or Triplicate)
11. Echocardiogram / Cardiac scintigraphy / MRI
12. Cuestionarios: KCCQ, EQ-5D-5L, SF-36, PGIC, Norfolk
13. Central clinical laboratory tests
14. sFLC, sIFE, and uIFE / Cardiac Biomarker Samples
15. Pharmacokinetic (PK) samples / Immunogenicity (ADA)
16. Adverse events / Concomitant Medications
17. Study intervention infusion / Study drug administration
"""

# --- 📄 GENERADOR DEL WORD DINÁMICO ---
def crear_documento_word_pro(datos_json, protocolo_nombre):
    try:
        match = re.search(r'\{.*\}', datos_json, re.DOTALL)
        datos = json.loads(match.group(0)) if match else json.loads(datos_json)
    except Exception as e:
        datos = {"visita": "Error", "procedimientos": {}, "detalles": {}}

    proc = datos.get("procedimientos", {})
    det = datos.get("detalles", {})
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(0, 0, 0)
    
    for i in range(1, 4):
        if f'Heading {i}' in doc.styles:
            h_style = doc.styles[f'Heading {i}']
            h_style.font.name = 'Calibri'
            h_style.font.bold = True
            h_style.font.color.rgb = RGBColor(0, 0, 0)

    es_alexion = "ALXN" in protocolo_nombre

    if es_alexion:
        # HOJA ENFERMERÍA ALEXION
        h_enf = doc.add_heading(f"HOJA DE ENFERMERÍA: {datos.get('visita', 'N/A')}", level=1)
        h_enf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Protocolo: {protocolo_nombre}     ID: ______________").bold = True
        
        doc.add_heading("PRIMEROS PASOS", level=2)
        doc.add_paragraph("☐ Registrar visita para asignar infusión en Almac (IXRS)")
        doc.add_paragraph("☐ Fecha de la visita: ______________")
        
        doc.add_heading("PRE-INFUSIÓN", level=2)
        doc.add_paragraph("☐ Peso: _________ kg")
        
        if proc.get("cuestionarios"):
            doc.add_paragraph("☐ Cuestionarios: KCCQ-OS, EQ-5D-5L, SF-36, PGIC").bold = True

        hdr = ["Tiempo", "Hora", "PA Sist.", "PA Diast.", "Pulso", "Resp.", "O2 (%)", "Ta (ºC)"]
        if proc.get("signos_vitales"):
            doc.add_paragraph("☐ Signos vitales: (reposo 5 min)").bold = True
            rows = 3 if proc.get("laboratorio") else 2
            t_vit = doc.add_table(rows=rows, cols=8)
            t_vit.style = 'Table Grid'
            for i, h in enumerate(hdr): t_vit.cell(0, i).text = h
            if proc.get("laboratorio"):
                t_vit.cell(1, 0).text = "Pre-extracción"
                t_vit.cell(2, 0).text = "Pre-infusión"
            else:
                t_vit.cell(1, 0).text = "Pre-infusión"

        if proc.get("ecg_pre"):
            doc.add_paragraph("☐ ECG 12 derivaciones Pre-infusión").bold = True

        if proc.get("laboratorio"):
            doc.add_paragraph("☐ Extracciones:").bold = True
            doc.add_paragraph(det.get("laboratorio_tubos", "Ver manual de laboratorio")).italic = True

        if proc.get("infusion"):
            doc.add_heading("INFUSIÓN", level=2)
            doc.add_paragraph("☐ HORA INICIO: _________   ☐ HORA FIN: _________")
            t_inf = doc.add_table(rows=5, cols=8)
            t_inf.style = 'Table Grid'
            for i, h in enumerate(hdr): t_inf.cell(0, i).text = h
            for r, t in enumerate(["15 min", "30 min", "45 min", "1h"]): t_inf.cell(r+1, 0).text = t

        # HOJA MÉDICA ALEXION
        doc.add_page_break()
        doc.add_heading(f"HOJA MÉDICA: {datos.get('visita', 'N/A')}", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        if proc.get("examen_fisico"): doc.add_paragraph("☐ Examen físico COMPLETO").bold = True
        if proc.get("examen_fisico_breve"): doc.add_paragraph("☐ Examen físico BREVE").bold = True
        if proc.get("nyha"): doc.add_paragraph("☐ NYHA: I [ ] II [ ] III [ ] IV [ ]").bold = True
        if proc.get("karnofsky"): doc.add_paragraph("☐ Karnofsky: ________ %").bold = True
        doc.add_paragraph("\nFirmado: _______________________")

    else:
        # PLANTILLA ALNYLAM 
        h_enf = doc.add_heading(f"{datos.get('visita', 'N/A')} - Hoja de Enfermería", level=1)
        h_enf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Protocol: {protocolo_nombre}     ID: ______________").bold = True
        
        if proc.get("laboratorio"):
            doc.add_paragraph("☐ Extraer muestras de sangre y orina").bold = True
            doc.add_paragraph(det.get("laboratorio_tubos", "Ver manual de laboratorio")).italic = True
            doc.add_paragraph("\n☐ Procesamiento de muestras (según manual)").bold = True
            doc.add_paragraph("☐ Envío de muestras").bold = True

        doc.add_page_break()
        h_med = doc.add_heading(f"{datos.get('visita', 'N/A')} - Hoja Médica", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("☐ Dar cita para próxima visita").bold = True
        doc.add_paragraph("☐ Medicamentos concomitantes").bold = True
        doc.add_paragraph("☐ Eventos adversos").bold = True
        doc.add_paragraph("☐ Registro Greenphire / Diraya / Medidata").bold = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def extraer_texto_paginas(pdf_bytes, paginas_str=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    try:
        p_list = []
        if paginas_str.strip():
            for b in paginas_str.split(','):
                if '-' in b:
                    s, e = b.split('-')
                    p_list.extend(range(int(s), int(e)+1))
                else: p_list.append(int(b))
        else: p_list = range(1, len(doc) + 1)
        
        for p in p_list:
            if 0 < p <= len(doc):
                page = doc[p-1]
                try: texto += f"\n--- PÁGINA {p} ---\n" + page.get_text("markdown") + "\n"
                except: texto += f"\n--- PÁGINA {p} ---\n" + page.get_text("text") + "\n"
    except Exception as e: 
        texto = f"Error lectura: {e}"
    return texto

# --- 🖥️ INTERFAZ WEB ---
st.sidebar.markdown("### 🔑 Configuración")
api_key = st.sidebar.text_input("Introduce API Key:", type="password")

# --- MODELOS ACTUALIZADOS SEGÚN TU LISTA ---
modelo_seleccionado = st.sidebar.selectbox("Modelo:", ["gemini-2.5-flash", "gemini-2.0-flash-lite", "gemini-3.1-flash-lite"])

st.markdown('<h1 style="color: #007d32;">BioCardio Clinical Assistant</h1>', unsafe_allow_html=True)
st.caption("v7.4 | Modelos actualizados + Optimización de tokens para evitar bloqueos")

st.markdown('<div class="step-header">📄 Paso 1: Documentación</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container">Sube los PDFs.</div>', unsafe_allow_html=True)
    f_proto = st.file_uploader("1. SUBIR PROTOCOLO (SoE)", type=["pdf"])
    f_labs = st.file_uploader("2. SUBIR MANUALES LAB", type=["pdf"], accept_multiple_files=True)

st.markdown('<div class="step-header">🔍 Paso 2: Configuración de la Visita</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container">Rellena las páginas. <b>¡Poner las páginas del manual de laboratorio evita que la web se bloquee!</b></div>', unsafe_allow_html=True)
    protocolo_sel = st.selectbox("Protocolo:", ["Alexion (ALXN2220-ATTR-CM-301)", "Alnylam (ALN-TTRSC04-003)"])
    
    c1, c2 = st.columns(2)
    p_tabla = c1.text_input("Páginas Tabla SoE (ej. 11-16):", "11-16")
    v_proto = c1.text_input("Visita Protocolo (ej. V25):", "V25")
    
    p_ass = c2.text_input("Páginas Assessments (ej. 40-60):", "40-60")
    v_lab = c2.text_input("Visita Manual Lab:", "Visita 25")
    
    # --- NUEVA CASILLA PARA EVITAR EL ERROR 429 ---
    p_lab = st.text_input("Páginas Manual Lab (Opcional, pero muy recomendado para evitar errores 429. Ej: 30-35):", "")

if st.button("✨ GENERAR HOJA OFICIAL"):
    if not api_key or not f_proto: st.error("Faltan datos.")
    else:
        with st.spinner("Analizando tablas y filtrando laboratorio..."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(modelo_seleccionado)
                p_bytes = f_proto.read()
                t_soe = extraer_texto_paginas(p_bytes, p_tabla)
                t_ass = extraer_texto_paginas(p_bytes, p_ass)
                
                t_lab = ""
                if f_labs:
                    for f in f_labs: 
                        # Ahora extrae solo las páginas indicadas del laboratorio
                        t_lab += extraer_texto_paginas(f.read(), p_lab)
                
                prompt = f"""
                Analiza la visita '{v_proto}' en esta TABLA SOE (MARKDOWN): {t_soe}.
                Busca la columna '{v_proto}'. Solo activa procedimientos con marca (X, *, punto).
                Si la visita no existe en la tabla, devuelve todo false.
                
                Extrae tubos de lab para '{v_lab}' de aquí: {t_lab[:50000]}.
                
                JSON:
                {{
                  "visita": "{v_proto}",
                  "procedimientos": {{
                    "cuestionarios": bool, "signos_vitales": bool, "ecg_pre": bool, "laboratorio": bool, "infusion": bool,
                    "examen_fisico": bool (si marca Full), "examen_fisico_breve": bool (si marca Symptom-directed),
                    "nyha": bool, "karnofsky": bool
                  }},
                  "detalles": {{ "laboratorio_tubos": "Lista de tubos extraída del texto..." }}
                }}
                """
                res = model.generate_content(prompt)
                st.download_button("⬇️ Descargar Word", crear_documento_word_pro(res.text, protocolo_sel), f"{v_proto}.docx")
            except Exception as e: st.error(f"Error: {e}")
