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

# =====================================================================
# 1. CONFIGURACIÓN E INTERFAZ (UI)
# =====================================================================
st.set_page_config(page_title="BioCardio Clinical Assistant", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f4; }
    .stFileUploader label { font-weight: bold !important; color: #1a1a1a !important; font-size: 19px !important; display: block; margin-bottom: 12px !important; }
    .step-container { background-color: white; padding: 25px 30px; border-radius: 15px; border-left: 8px solid #007d32; box-shadow: 0 4px 12px rgba(0,0,0,0.08); margin-bottom: 35px; }
    .step-header { color: #007d32; font-size: 24px; font-weight: bold; margin-top: 15px; margin-bottom: 10px; }
    .stButton>button { background-color: #007d32 !important; color: white !important; border-radius: 12px !important; height: 55px !important; font-size: 18px !important; font-weight: bold !important; width: 100%; transition: 0.3s; border: none !important; }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# 2. MOTOR DE EXTRACCIÓN DE PDF (MARKDOWN PARA TABLAS)
# =====================================================================
def extraer_texto_pdf(pdf_bytes, paginas_str=""):
    """ Extrae el texto de un PDF. Si se indican páginas, usa formato Markdown para no romper las tablas. """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    try:
        # Calcular qué páginas leer
        p_list = []
        if paginas_str.strip():
            for b in paginas_str.split(','):
                if '-' in b:
                    s, e = b.split('-')
                    p_list.extend(range(int(s), int(e)+1))
                else: 
                    p_list.append(int(b))
        else:
            p_list = range(1, len(doc) + 1)
        
        # Extraer página a página
        for p in p_list:
            if 0 < p <= len(doc):
                page = doc[p-1]
                texto += f"\n--- PÁGINA {p} ---\n"
                # Usar markdown para conservar la estructura de columnas | V24 | V25 |
                try: 
                    texto += page.get_text("markdown") + "\n"
                except: 
                    texto += page.get_text("text") + "\n"
    except Exception as e:
        texto = f"Error al leer PDF: {e}"
    return texto

# =====================================================================
# 3. GENERADOR DEL DOCUMENTO WORD (ESTILOS Y PLANTILLAS)
# =====================================================================
def crear_documento_word(datos_json, protocolo_nombre):
    # Intentar parsear la respuesta de la IA
    try:
        match = re.search(r'\{.*\}', datos_json, re.DOTALL)
        datos = json.loads(match.group(0)) if match else json.loads(datos_json)
    except:
        datos = {"visita": "Error", "procedimientos": {}, "detalles": {}}

    proc = datos.get("procedimientos", {})
    det = datos.get("detalles", {})
    doc = Document()
    
    # FORZAR TODO A NEGRO Y TÍTULOS EN NEGRITA
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

    # ---------------------------------------------------------
    # PLANTILLA A: ALEXION (ALXN2220-ATTR-CM-301)
    # ---------------------------------------------------------
    if es_alexion:
        # --- PÁGINA 1: ENFERMERÍA ---
        doc.add_heading(f"HOJA DE ENFERMERÍA: {datos.get('visita', 'N/A')}", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Protocolo: {protocolo_nombre}     ID: ______________").bold = True
        
        doc.add_heading("PRIMEROS PASOS", level=2)
        doc.add_paragraph("☐ Registrar visita para asignar infusión en Almac (IXRS)")
        doc.add_paragraph("☐ Fecha de la visita: ______________")
        
        doc.add_heading("PRE-INFUSIÓN (dentro de las 2 horas antes de la infusión)", level=2)
        doc.add_paragraph("☐ Peso: _________ kg")
        
        if proc.get("cuestionarios"):
            doc.add_paragraph("☐ Cuestionarios o PROs (TABLET en armario):").bold = True
            doc.add_paragraph("    ☐ KCCQ-OS  ☐ EQ-5D-5L  ☐ SF-36  ☐ PGIC (EN PAPEL)")

        hdr = ["Tiempo", "Hora", "PA Sist.", "PA Diast.", "Pulso", "Resp.", "O2 (%)", "Ta (ºC)"]
        if proc.get("signos_vitales"):
            doc.add_paragraph("☐ Signos vitales: (tras 5 minutos descanso)").bold = True
            rows = 3 if proc.get("laboratorio") else 2
            t_vit_pre = doc.add_table(rows=rows, cols=8)
            t_vit_pre.style = 'Table Grid'
            for i, h in enumerate(hdr): t_vit_pre.cell(0, i).text = h
            if proc.get("laboratorio"):
                t_vit_pre.cell(1, 0).text = "Pre-extracción"
                t_vit_pre.cell(2, 0).text = "Pre-infusión"
            else:
                t_vit_pre.cell(1, 0).text = "Pre-infusión"
            doc.add_paragraph("")

        if proc.get("ecg_pre"):
            doc.add_paragraph("☐ Pre-Electrocardiograma de 12 derivaciones").bold = True
            doc.add_paragraph("    ☐ Posición supino    ☐ FC: ____  ☐ PR: ____  ☐ QRS: ____  ☐ QT: ____  ☐ QTc: ____")

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

        # --- PÁGINA 2: MÉDICO ---
        doc.add_page_break()
        doc.add_heading(f"HOJA MÉDICA: {datos.get('visita', 'N/A')}", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Protocolo: {protocolo_nombre}     ID: ______________").bold = True
        doc.add_paragraph("☐ Fecha de la visita: ______________\n")
        
        if proc.get("examen_fisico"): doc.add_paragraph("☐ Examen físico COMPLETO").bold = True
        if proc.get("examen_fisico_breve"): doc.add_paragraph("☐ Examen físico BREVE (Dirigido por síntomas)").bold = True
        if proc.get("nyha"): doc.add_paragraph("☐ Clasificación NYHA:  I [ ]   II [ ]   III [ ]   IV [ ]").bold = True
        if proc.get("karnofsky"): doc.add_paragraph("☐ Discapacidad (Karnofsky): 100% a 0% evaluado.").bold = True
        
        doc.add_paragraph("☐ Medicamentos, terapias o procedimientos simultáneos").bold = True
        doc.add_paragraph("☐ Eventos adversos (AEs)\n").bold = True
        doc.add_paragraph("\nFirmado (Médico): _______________________      Fecha: ______________")

    # ---------------------------------------------------------
    # PLANTILLA B: ALNYLAM (ALN-TTRSC04-003)
    # ---------------------------------------------------------
    else:
        # --- PÁGINA 1: ENFERMERÍA ---
        doc.add_heading(f"{datos.get('visita', 'N/A')} - Hoja de Enfermería", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
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
            doc.add_paragraph("\n☐ Procesamiento de muestras en el laboratorio (según manual)").bold = True
            doc.add_paragraph("☐ Envío de muestras a Tª ambiente y congeladas").bold = True

        if proc.get("infusion"):
            doc.add_paragraph("\n☐ Administración de medicación del estudio (Study drug administration)").bold = True

        doc.add_paragraph("\n\nFirma: ______________________________        Fecha: ______________")

        # --- PÁGINA 2: MÉDICO ---
        doc.add_page_break()
        doc.add_heading(f"{datos.get('visita', 'N/A')} - General / Médica", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("______________________________________________________________")
        doc.add_paragraph(f"Protocol: {protocolo_nombre}\nID:\nFecha de la visita:").bold = True
        doc.add_paragraph("______________________________________________________________\n")
        
        doc.add_paragraph("☐ Dar cita para próxima visita: ___________________").bold = True
        doc.add_paragraph("\n☐ Medicamentos concomitantes, terapias o procedimientos actuales:").bold = True
        doc.add_paragraph("\n☐ Eventos adversos desde la firma del consentimiento:").bold = True

        if proc.get("examen_fisico") or proc.get("examen_fisico_breve"): 
            doc.add_paragraph("\n☐ Examen físico (Completo o dirigido por síntomas)").bold = True
        if proc.get("nyha"): doc.add_paragraph("☐ Clasificación NYHA evaluada").bold = True
        if proc.get("karnofsky"): doc.add_paragraph("☐ Puntuación Karnofsky (Karnofsky Performance Status)").bold = True
        if proc.get("test_6mwt"): doc.add_paragraph("☐ Test de los 6 minutos (6MWT) completado").bold = True

        doc.add_paragraph("\nCIERRE DE VISITA:").bold = True
        doc.add_paragraph("☐ Registrar próxima visita en Greenphire")
        doc.add_paragraph("☐ Escribir historia clínica en Diraya (imprimir, fechar y firmar)")
        doc.add_paragraph("☐ Rellenar CRF (Medidata)")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =====================================================================
# 4. APLICACIÓN STREAMLIT (LÓGICA PRINCIPAL)
# =====================================================================
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try: st.image("huj.png", width=140)
    except: st.title("🏥")
with col_titulo:
    st.markdown("<h1 style='color: #007d32; margin-top: 10px;'>BioCardio Clinical Assistant</h1>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 18px; color: #555;'>Herramienta profesional de gestión de ensayos clínicos</p>", unsafe_allow_html=True)
    st.caption("v9.0 | Limpieza Total de Código + Extracción Markdown")

with st.sidebar:
    st.markdown("### 🔑 Configuración")
    api_key = st.text_input("Introduce API Key:", type="password")
    modelo_seleccionado = st.selectbox("Modelo:", ["gemini-2.5-flash", "gemini-2.0-flash-lite"])

st.markdown('<div class="step-header">📄 Paso 1: Documentación</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container">Sube el protocolo (SoE) y los manuales de laboratorio íntegros.</div>', unsafe_allow_html=True)
    f_proto = st.file_uploader("1. SUBIR PROTOCOLO (SoE)", type=["pdf"])
    f_labs = st.file_uploader("2. SUBIR MANUALES LAB", type=["pdf"], accept_multiple_files=True)

st.markdown('<div class="step-header">🔍 Paso 2: Configuración de la Visita</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container">Configura los parámetros para crear la plantilla.</div>', unsafe_allow_html=True)
    protocolo_sel = st.selectbox("Protocolo a maquetar:", ["Alexion (ALXN2220-ATTR-CM-301)", "Alnylam (ALN-TTRSC04-003)"])
    
    c1, c2 = st.columns(2)
    p_tabla = c1.text_input("Páginas Tabla SoE (ej. 11-16):", "11-16")
    v_proto = c1.text_input("Visita Protocolo (ej. V25b):", "V25b")
    p_ass = c2.text_input("Páginas Assessments (ej. 40-60):", "40-60")
    v_lab = c2.text_input("Visita Manual Lab:", "Visita 25")

st.markdown('<div class="step-header">📋 Paso 3: Generación del Checklist</div>', unsafe_allow_html=True)
with st.container():
    if st.button("✨ GENERAR HOJA OFICIAL"):
        if not api_key or not f_proto: 
            st.error("⚠️ Faltan documentos o la API Key.")
        else:
            barra_progreso = st.progress(0)
            status_text = st.empty()
            
            try:
                # 1. Extraer Protocolo
                status_text.text("📖 Analizando tabla del protocolo en formato columnas...")
                p_bytes = f_proto.read()
                t_soe = extraer_texto_pdf(p_bytes, p_tabla)
                t_ass = extraer_texto_pdf(p_bytes, p_ass)
                barra_progreso.progress(40)
                
                # 2. Extraer Manuales Lab (Completos, sin pedir páginas)
                status_text.text("🧪 Analizando el manual de laboratorio completo...")
                t_lab = ""
                if f_labs:
                    for f in f_labs: 
                        t_lab += extraer_texto_pdf(f.read(), "")
                barra_progreso.progress(70)
                
                # 3. Inteligencia Artificial (Cerebro)
                status_text.text("🧠 IA evaluando qué procedimientos corresponden...")
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(modelo_seleccionado)
                
                prompt = f"""
                Eres un Data Manager Clínico. Tu objetivo es leer una tabla en formato Markdown y un manual de laboratorio.
                
                --- TABLA SOE (MARKDOWN) --- 
                {t_soe}
                
                --- DETALLES TÉCNICOS --- 
                {t_ass}
                
                --- LABORATORIO --- 
                {t_lab[:80000]}

                INSTRUCCIONES CRÍTICAS (TOLERANCIA CERO A INVENTAR):
                1. Busca la columna exacta '{v_proto}' en la tabla SoE.
                2. Si la columna no existe, devuelve todos los booleanos a FALSE.
                3. Si existe, baja por esa columna y pon TRUE solo a los procedimientos que tengan una marca en esa columna específica.
                4. LABORATORIO: Si hay extracciones, busca la visita '{v_lab}' en el manual de laboratorio adjunto y resume los tubos, volúmenes y colores.
                
                FORMATO JSON DE SALIDA OBLIGATORIO:
                {{
                  "visita": "{v_proto}",
                  "procedimientos": {{
                    "cuestionarios": bool, "signos_vitales": bool, "ecg_pre": bool, "ecg_post": bool, "laboratorio": bool, "infusion": bool,
                    "examen_fisico": bool, "examen_fisico_breve": bool,
                    "nyha": bool, "karnofsky": bool, "test_6mwt": bool, "pk_post": bool, "eco": bool
                  }},
                  "detalles": {{ "laboratorio_tubos": "Resumen de tubos y colores..." }}
                }}
                """
                res = model.generate_content(prompt)
                
                # 4. Generar el Word
                status_text.text("📝 Maquetando documento Word...")
                doc_word = crear_documento_word(res.text, protocolo_sel)
                
                barra_progreso.progress(100)
                status_text.success("✅ Documento generado con éxito.")
                st.download_button("⬇️ Descargar Documento Word", doc_word, f"{v_proto}.docx")
                
            except Exception as e: 
                st.error(f"Error crítico en el proceso: {e}")
