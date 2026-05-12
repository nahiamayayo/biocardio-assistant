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

# --- 📄 GENERADOR DEL WORD DINÁMICO (ALEXION vs ALNYLAM) ---
def crear_documento_word_pro(datos_json, protocolo_nombre):
    try:
        match = re.search(r'\{.*\}', datos_json, re.DOTALL)
        datos = json.loads(match.group(0)) if match else json.loads(datos_json)
    except Exception as e:
        datos = {"visita": "Error", "procedimientos": {}, "detalles": {}}

    proc = datos.get("procedimientos", {})
    det = datos.get("detalles", {})
    doc = Document()
    
    # === FORZAR ESTILOS A NEGRO Y TÍTULOS EN NEGRITA ===
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

    # =========================================================================
    # PLANTILLA 1: ALEXION
    # =========================================================================
    if es_alexion:
        h_enf = doc.add_heading(f"HOJA DE ENFERMERÍA: {datos.get('visita', 'N/A')}", level=1)
        h_enf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Protocolo: {protocolo_nombre}     ID: ______________").bold = True
        
        doc.add_heading("PRIMEROS PASOS", level=2)
        doc.add_paragraph("☐ Registrar visita para asignar infusión al paciente en Almac (IXRS)")
        doc.add_paragraph("☐ Fecha de la visita: ______________")
        
        doc.add_heading("PRE-INFUSIÓN (dentro de las 2 horas antes de la infusión)", level=2)
        doc.add_paragraph("☐ Peso: _________ kg")
        
        if proc.get("cuestionarios"):
            doc.add_paragraph("☐ Cuestionarios o PROs (TABLET en armario):").bold = True
            doc.add_paragraph("    ☐ KCCQ-OS  ☐ EQ-5D-5L  ☐ SF-36  ☐ PGIC (EN PAPEL)")

        hdr = ["Tiempo", "Hora", "PA Sist.", "PA Diast.", "Pulso", "Resp.", "O2 (%)", "Ta (ºC)"]
        
        # --- LÓGICA CORREGIDA DE SIGNOS VITALES ---
        if proc.get("signos_vitales"):
            doc.add_paragraph("☐ Signos vitales: (tras 5 minutos descanso)").bold = True
            if proc.get("laboratorio"):
                # Hay laboratorio: Doble toma
                t_vit_pre = doc.add_table(rows=3, cols=8)
                t_vit_pre.style = 'Table Grid'
                for i, h in enumerate(hdr): t_vit_pre.cell(0, i).text = h
                t_vit_pre.cell(1, 0).text = "Pre-extracción"
                t_vit_pre.cell(2, 0).text = "Pre-infusión"
            else:
                # No hay laboratorio: Toma simple
                t_vit_pre = doc.add_table(rows=2, cols=8)
                t_vit_pre.style = 'Table Grid'
                for i, h in enumerate(hdr): t_vit_pre.cell(0, i).text = h
                t_vit_pre.cell(1, 0).text = "Pre-infusión"
            doc.add_paragraph("")

        if proc.get("ecg_pre"):
            doc.add_paragraph("☐ Pre-Electrocardiograma de 12 derivaciones").bold = True
            doc.add_paragraph("    ☐ Posición supino\n    ☐ FC: ____  ☐ PR: ____  ☐ QRS: ____  ☐ QT: ____  ☐ QTc: ____")

        if proc.get("laboratorio"):
            doc.add_paragraph("☐ Extraer muestras de sangre y orina").bold = True
            doc.add_paragraph(f"Tubos a extraer según manual: {det.get('laboratorio_tubos', 'Ver manual')}").italic = True

        if proc.get("test_6mwt"):
            doc.add_paragraph("☐ Test de los 6 minutos (6MWT)").bold = True
            doc.add_paragraph("    • Rellenar en Course Name: 6MWT 2F; y en Course Lenght: 20m")

        if proc.get("infusion"):
            doc.add_heading("ADMINISTRACIÓN DE INFUSIÓN + SIGNOS VITALES", level=2)
            doc.add_paragraph("*La infusión debe durar 1 hora. Limpiar vías con dextrosa 5%.").italic = True
            doc.add_paragraph("☐ HORA DE INICIO: _________   ☐ HORA DE FIN: _________")
            doc.add_paragraph("\nSignos vitales durante la infusión:").bold = True
            t_vit_inf = doc.add_table(rows=5, cols=8)
            t_vit_inf.style = 'Table Grid'
            for i, h in enumerate(hdr): t_vit_inf.cell(0, i).text = h
            for r, text in enumerate(["15 min infusión", "30 min infusión", "45 min infusión", "1h infusión"]):
                t_vit_inf.cell(r+1, 0).text = text
            doc.add_paragraph("")

        doc.add_heading("POST-INFUSIÓN (30 min)", level=2)
        if proc.get("pk_post"): doc.add_paragraph("☐ PK-post infusión: sacar del brazo opuesto.").bold = True
        if proc.get("ecg_post"): doc.add_paragraph("☐ Post-Electrocardiograma.").bold = True
        if proc.get("signos_vitales"):
            t_vit_post = doc.add_table(rows=2, cols=8)
            t_vit_post.style = 'Table Grid'
            for i, h in enumerate(hdr): t_vit_post.cell(0, i).text = h
            t_vit_post.cell(1, 0).text = "30 min post"
        
        doc.add_paragraph("\nFirmado: _______________________      Fecha: ______________")

        # PÁGINA 2: HOJA MÉDICA (Alexion)
        doc.add_page_break()
        h_med = doc.add_heading(f"HOJA MÉDICA: {datos.get('visita', 'N/A')}", level=1)
        h_med.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Protocolo: {protocolo_nombre}     ID: ______________").bold = True
        doc.add_paragraph("☐ Fecha de la visita: ______________\n")
        
        # --- LÓGICA CORREGIDA DE EXAMEN FÍSICO ---
        if proc.get("examen_fisico"): 
            doc.add_paragraph("☐ Examen físico COMPLETO").bold = True
            doc.add_paragraph("(Ver aspecto general, piel, nariz, orejas, ojos, cuello, garganta, corazón, abdomen, pulmones, sistema vascular, sistema nervioso, sistema musculo esquelético y extremidades)\n")
        if proc.get("examen_fisico_breve"): 
            doc.add_paragraph("☐ Examen físico BREVE (Dirigido por síntomas)").bold = True
            doc.add_paragraph("Inclusive of general appearance, heart, lungs, skin, musculoskeletal system and extremities and other organs or body systems as clinically indicated.\n")
        
        if proc.get("nyha"): doc.add_paragraph("☐ Clasificación NYHA:  I [ ]   II [ ]   III [ ]   IV [ ]").bold = True
        if proc.get("karnofsky"): doc.add_paragraph("☐ Discapacidad (Karnofsky): 100% a 0% evaluado.").bold = True
        doc.add_paragraph("☐ Medicamentos, terapias o procedimientos simultáneos").bold = True
        doc.add_paragraph("☐ Eventos adversos (AEs)\n").bold = True
        doc.add_paragraph("\nFirmado (Médico): _______________________      Fecha: ______________")

    # =========================================================================
    # PLANTILLA 2: ALNYLAM
    # =========================================================================
    else:
        h_enf = doc.add_heading(f"{datos.get('visita', 'N/A')} - Hoja de Enfermería", level=1)
        h_enf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("______________________________________________________________")
        doc.add_paragraph(f"Protocol: {protocolo_nombre}\nID:\nFecha de la visita:").bold = True
        doc.add_paragraph("______________________________________________________________\n")
        
        if proc.get("signos_vitales"):
            doc.add_paragraph("☐ Signos vitales:").bold = True
            hdr = ["Tiempo", "Hora", "PA Sist.", "PA Diast.", "Pulso", "Resp.", "O2 (%)", "Ta (ºC)"]
            t_vit = doc.add_table(rows=2, cols=8)
            t_vit.style = 'Table Grid'
            for i, h in enumerate(hdr): t_vit.cell(0, i).text = h
            doc.add_paragraph("")

        if proc.get("ecg_pre"):
            doc.add_paragraph("☐ Electrocardiograma de 12 derivaciones").bold = True
            doc.add_paragraph("    ☐ FC: ____  ☐ PR: ____  ☐ QRS: ____  ☐ QT: ____  ☐ QTc: ____")

        if proc.get("cuestionarios"):
            doc.add_paragraph("☐ Cuestionarios y PROs realizados.").bold = True
            
        if proc.get("laboratorio"):
            doc.add_paragraph("☐ Extraer muestras de sangre y orina para analítica central").bold = True
            doc.add_paragraph(f"{det.get('laboratorio_tubos', 'Ver manual de laboratorio')}").italic = True
            doc.add_paragraph("\n☐ Procesamiento de muestras en el laboratorio").bold = True
            doc.add_paragraph("*(Completar tiras reactivas, centrifugado y alícuotas según el manual del estudio)*").italic = True
            doc.add_paragraph("\n☐ Envío de muestras a Tª ambiente y congeladas").bold = True

        if proc.get("infusion"):
            doc.add_paragraph("\n☐ Administración de medicación del estudio (Study drug administration)").bold = True

        doc.add_paragraph("\n\nFirma: ______________________________        Fecha: ______________")

        # PÁGINA 2: HOJA MÉDICA / GENERAL (Alnylam)
        doc.add_page_break()
        h_med = doc.add_heading(f"{datos.get('visita', 'N/A')} - General / Médica", level=1)
        h_med.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("______________________________________________________________")
        doc.add_paragraph(f"Protocol: {protocolo_nombre}\nID:\nFecha de la visita:").bold = True
        doc.add_paragraph("______________________________________________________________\n")
        
        doc.add_paragraph("☐ Dar cita para próxima visita: ___________________").bold = True
        doc.add_paragraph("\n☐ Medicamentos concomitantes, terapias o procedimientos actuales:").bold = True
        doc.add_paragraph("    □ ........................................................................")
        doc.add_paragraph("\n☐ Eventos adversos desde la firma del consentimiento:").bold = True
        doc.add_paragraph("    □ ........................................................................")

        if proc.get("examen_fisico") or proc.get("examen_fisico_breve"): 
            doc.add_paragraph("\n☐ Examen físico (Completo o dirigido por síntomas)").bold = True
        if proc.get("nyha"): doc.add_paragraph("☐ Clasificación NYHA evaluada").bold = True
        if proc.get("karnofsky"): doc.add_paragraph("☐ Puntuación Karnofsky (Karnofsky Performance Status)").bold = True
        if proc.get("test_6mwt"): doc.add_paragraph("☐ Test de los 6 minutos (6MWT) completado").bold = True

        doc.add_paragraph("\nCIERRE DE VISITA:").bold = True
        doc.add_paragraph("☐ Registrar próxima visita en Greenphire")
        doc.add_paragraph("☐ Escribir historia clínica en Diraya y debe imprimirse, fecharse, firmarse y guardar en físico")
        doc.add_paragraph("☐ Rellenar CRF (Medidata)")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- 🔍 EXTRACCIÓN DE TEXTO ---
def extraer_texto_paginas(pdf_bytes, paginas_str=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    try:
        paginas_a_leer = []
        if paginas_str.strip(): 
            p_set = set()
            for b in paginas_str.split(','):
                if '-' in b:
                    s, e = b.split('-')
                    p_set.update(range(int(s), int(e)+1))
                else: 
                    p_set.add(int(b))
            paginas_a_leer = sorted(list(p_set))
        else:
            paginas_a_leer = range(1, len(doc) + 1)
            
        for p in paginas_a_leer:
            if 0 < p <= len(doc):
                texto += f"\n--- PÁGINA {p} ---\n" + doc[p-1].get_text("text") + "\n"
    except Exception as e: 
        texto = f"Error en lectura de PDF: {e}"
    return texto

# --- 🖥️ INTERFAZ WEB ---
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try: st.image("huj.png", width=140)
    except: st.title("🏥")
with col_titulo:
    st.markdown("<h1 style='color: #007d32; margin-top: 10px;'>BioCardio Clinical Assistant</h1>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 18px; color: #555;'>Herramienta profesional de gestión de ensayos clínicos</p>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🔑 Acceso")
    st.markdown("Obtén tu clave en [Google AI Studio ↗](https://aistudio.google.com/app/apikey)")
    api_key = st.text_input("Introduce tu API Key:", type="password")
    st.divider()
    modelo_seleccionado = st.selectbox("Modelo de IA:", ["gemini-3.1-flash-lite-preview"])
    st.caption("v7.1 | Lógica Condicional + Anti-Alucinaciones")

st.markdown('<div class="step-header">📄 Paso 1: Documentación</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">Carga el Protocolo SoE (Las X) y los manuales de lab.</div>', unsafe_allow_html=True)
    f_proto = st.file_uploader("1. SUBIR PROTOCOLO (Apartado SoE)", type=["pdf"])
    f_labs = st.file_uploader("2. SUBIR MANUALES DE LABORATORIO", type=["pdf"], accept_multiple_files=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="step-header">🔍 Paso 2: Configuración de la Visita</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">Indica el protocolo, las páginas y la visita a generar.</div>', unsafe_allow_html=True)
    protocolo_sel = st.selectbox("Selecciona el Protocolo a maquetar:", ["Alexion (ALXN2220-ATTR-CM-301)", "Alnylam (ALN-TTRSC04-003)"])
    
    c1, c2 = st.columns(2)
    p_tabla = c1.text_input("Páginas de la Tabla SoE (ej. 11-16):", "11-16")
    v_proto = c1.text_input("Visita en Protocolo (ej. Visit 17):", "Visit 17")
    p_assessments = c2.text_input("Páginas Study Assessments:", "40-60")
    v_lab = c2.text_input("Visita en Manual Lab:", "Visit 17")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="step-header">📋 Paso 3: Generación del Checklist</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">La IA analizará el protocolo y generará las Hojas oficiales según el diseño del ensayo seleccionado.</div>', unsafe_allow_html=True)
    
    if st.button("✨ Generar Hoja de Visita Oficial HUJ"):
        if not api_key or not f_proto:
            st.error("⚠️ Faltan documentos o la API Key.")
        else:
            barra_progreso = st.progress(0)
            status_text = st.empty()
            
            try:
                status_text.text("📖 Leyendo la tabla del protocolo y los assessments...")
                p_bytes = f_proto.read()
                t_soe = extraer_texto_paginas(p_bytes, p_tabla)
                t_ass = extraer_texto_paginas(p_bytes, p_assessments)
                barra_progreso.progress(30)

                status_text.text("🧪 Analizando manuales de laboratorio...")
                t_lab = ""
                if f_labs:
                    for f in f_labs:
                        t_lab += f"\n--- DOC: {f.name} ---\n" + extraer_texto_paginas(f.read(), "")
                barra_progreso.progress(60)
                
                status_text.text("🧠 IA cruzando información (Modo Estricto)...")
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(modelo_seleccionado)
                
                prompt = f"""
                Eres el Director Médico del ensayo. Analiza la tabla de protocolo proporcionada para la visita '{v_proto}'.
                
                --- LISTA MAESTRA DE PROCEDIMIENTOS POSIBLES ---
                {PROCEDIMIENTOS_MAESTROS}
                
                --- TABLA SOE --- 
                {t_soe}
                
                --- DETALLES TÉCNICOS --- 
                {t_ass}
                
                --- LABORATORIO --- 
                {t_lab[:80000]}

                REGLA DE TOLERANCIA CERO A LAS ALUCINACIONES:
                1. Busca la columna EXACTA de la visita '{v_proto}' en el texto de la TABLA SOE.
                2. Si la visita '{v_proto}' NO ESTÁ en el texto de la tabla, pon TODOS los booleanos a false. NO inventes nada.
                3. Solo pon "true" si hay una marca explícita (X, asterisco, punto) en la intersección del procedimiento y la columna de la visita.

                REGLAS DE MAPEO (Si la visita sí existe):
                - examen_fisico: true SOLO si marca 'Full physical examination'.
                - examen_fisico_breve: true SOLO si marca 'Symptom-directed physical examination' o examen breve.
                - laboratorio: true si hay marcas en Central clinical laboratory tests o extracciones.
                - LABORATORIO_TUBOS: Si 'laboratorio' es true, busca '{v_lab}' en el manual de laboratorio y extrae los tubos, volúmenes y colores.

                Formato estricto de salida:
                {{
                  "visita": "{v_proto}",
                  "procedimientos": {{
                    "cuestionarios": true|false,
                    "signos_vitales": true|false,
                    "ecg_pre": true|false,
                    "ecg_post": true|false,
                    "test_6mwt": true|false,
                    "laboratorio": true|false,
                    "infusion": true|false,
                    "pk_post": true|false,
                    "eco": true|false,
                    "examen_fisico": true|false,
                    "examen_fisico_breve": true|false,
                    "nyha": true|false,
                    "karnofsky": true|false
                  }},
                  "detalles": {{
                    "laboratorio_tubos": "Resumen exacto de los tubos y colores a extraer...",
                    "instrucciones_ecg": "Detalles técnicos del ECG...",
                    "instrucciones_6mwt": "Detalles técnicos del test de marcha..."
                  }}
                }}
                """
                
                res = model.generate_content(prompt)
                doc_word = crear_documento_word_pro(res.text, protocolo_sel)
                
                barra_progreso.progress(100)
                status_text.success("✅ Documentos médicos oficiales generados con éxito.")
                
                st.download_button(
                    label="⬇️ Descargar Documento Word",
                    data=doc_word,
                    file_name=f"Hojas_Oficiales_{v_proto.replace(' ','_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error técnico: {e}")
                
    st.markdown("</div>", unsafe_allow_html=True)
