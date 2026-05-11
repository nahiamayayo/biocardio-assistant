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

# --- 📄 GENERADOR DEL WORD (PLANTILLA EXACTA HUJ) ---
def crear_documento_word_pro(datos_json):
    try:
        match = re.search(r'\{.*\}', datos_json, re.DOTALL)
        datos = json.loads(match.group(0)) if match else json.loads(datos_json)
    except Exception as e:
        datos = {"visita": "Error", "procedimientos": {}}

    proc = datos.get("procedimientos", {})
    doc = Document()
    
    # Estilos Base
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    
    # ==========================================
    # PÁGINA 1: HOJA DE ENFERMERÍA
    # ==========================================
    doc.add_heading(f"HOJA DE ENFERMERÍA: {datos.get('visita', 'N/A')}", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Protocolo: {datos.get('protocolo', 'ALXN2220-ATTR-CM-301')}     ID: ______________").bold = True
    
    doc.add_heading("PRIMEROS PASOS", level=2)
    doc.add_paragraph("☐ Registrar visita para asignar infusión al paciente en Almac (IXRS)")
    doc.add_paragraph("☐ Fecha de la visita: ______________")
    doc.add_paragraph("*Si no viene a la visita también rellenar CRF\n").italic = True
    
    doc.add_heading("PRE-INFUSIÓN (dentro de las 2 horas antes de la infusión)", level=2)
    doc.add_paragraph("☐ Peso: _________ kg")
    
    # Cuestionarios
    if proc.get("cuestionarios"):
        doc.add_paragraph("☐ Cuestionarios o PROs (TABLET en armario):").bold = True
        doc.add_paragraph("    ☐ KCCQ-OS\n    ☐ EQ-5D-5L\n    ☐ SF-36\n    ☐ PGIC (EN PAPEL)")

    # Signos Vitales Pre
    if proc.get("signos_vitales"):
        doc.add_paragraph("☐ Signos vitales: (antes de sacar la sangre, tras 5 minutos descanso)").bold = True
        doc.add_paragraph("Medir las constantes en posición de reposo y tras haber estado 5 minutos en reposo.")
        t_vit_pre = doc.add_table(rows=2, cols=8)
        t_vit_pre.style = 'Table Grid'
        hdr = ["Tiempo", "Hora", "PA Sist.", "PA Diast.", "Pulso", "Resp.", "O2 (%)", "Ta (ºC)"]
        for i, h in enumerate(hdr): t_vit_pre.cell(0, i).text = h
        t_vit_pre.cell(1, 0).text = "Pre-extracción"
        doc.add_paragraph("")

    # ECG Pre
    if proc.get("ecg_pre"):
        doc.add_paragraph("☐ Pre-Electrocardiograma de 12 derivaciones (dentro de las 2h antes)").bold = True
        doc.add_paragraph("Hacer el ECG después de haber estado en reposo 5 minutos.")
        doc.add_paragraph("    ☐ Posición supino\n    ☐ FC: ____  ☐ PR: ____  ☐ QRS: ____  ☐ QT: ____  ☐ QTc: ____")

    # Analíticas
    if proc.get("laboratorio"):
        doc.add_paragraph("☐ Extraer muestras de sangre y orina (dentro de las 2h antes)").bold = True
        doc.add_paragraph("5 tubos rojos, 3 dorados, 1 azul y 2 lavanda (EL TUBO ROJO PK EOI ES POST INFUSIÓN)")
        doc.add_paragraph("☐ Orina")

    # 6MWT
    if proc.get("test_6mwt"):
        doc.add_paragraph("☐ Test de los 6 minutos (6MWT) (después de sangre, antes de infusión)").bold = True
        doc.add_paragraph("    • Coger material y hoja, escribir en inglés. Realizado por: ________________")
        doc.add_paragraph("    • Rellenar en Course Name: 6MWT 2F; y en Course Lenght: 20m")
        doc.add_paragraph("    • Todas las casillas completas. Si no hiciste alguna: N/D.")

    # INFUSIÓN
    if proc.get("infusion"):
        doc.add_heading("ADMINISTRACIÓN DE INFUSIÓN + SIGNOS VITALES", level=2)
        doc.add_paragraph("*La infusión debe durar 1 hora.\n*No diluir el fármaco\n*Limpiar las vías ANTES y DESPUÉS con dextrosa 5%.\n*Completar el IMP Transfer Log.").italic = True
        doc.add_paragraph("Administración del fármaco:").bold = True
        doc.add_paragraph("1. Revisar visualmente la bolsa. NO SACAR la bolsa de cegado.\n2. Revisar caducidad.\n3. Administrar con filtro de línea de 0.2 micras.\n4. Seguir velocidad de la tabla.\n5. Registrar hora de inicio y brazo. Signos vitales cada 15 min.\n6. Limpiar vía con dextrosa 5% (5 min).\n7. Quitar bolsa.")
        doc.add_paragraph("\n☐ HORA DE INICIO: _________   ☐ HORA DE FIN: _________")
        doc.add_paragraph("☐ Rellenar IMP transfer log y revisar bolsa de cegado.")
        
        doc.add_paragraph("\nSignos vitales durante la infusión:").bold = True
        t_vit_inf = doc.add_table(rows=6, cols=8)
        t_vit_inf.style = 'Table Grid'
        for i, h in enumerate(hdr): t_vit_inf.cell(0, i).text = h
        t_vit_inf.cell(1, 0).text = "15 min antes"
        t_vit_inf.cell(2, 0).text = "15 min infusión"
        t_vit_inf.cell(3, 0).text = "30 min infusión"
        t_vit_inf.cell(4, 0).text = "45 min infusión"
        t_vit_inf.cell(5, 0).text = "1h infusión"
        doc.add_paragraph("")

    # POST-INFUSIÓN
    doc.add_heading("POST-INFUSIÓN (30 min)", level=2)
    if proc.get("pk_post"):
        doc.add_paragraph("☐ PK-post infusión: dentro de los 30 min post-infusión.").bold = True
        doc.add_paragraph("Sacar la muestra del brazo opuesto de la infusión.")
    
    if proc.get("ecg_post"):
        doc.add_paragraph("☐ Post-Electrocardiograma: dentro de los 30 min post-infusion y 5m reposo.").bold = True
        doc.add_paragraph("    ☐ Posición supino\n    ☐ FC: ____  ☐ PR: ____  ☐ QRS: ____  ☐ QT: ____  ☐ QTc: ____")
        
    if proc.get("signos_vitales"):
        t_vit_post = doc.add_table(rows=2, cols=8)
        t_vit_post.style = 'Table Grid'
        for i, h in enumerate(hdr): t_vit_post.cell(0, i).text = h
        t_vit_post.cell(1, 0).text = "30 min post"
        doc.add_paragraph("")

    if proc.get("eco"):
        doc.add_paragraph("☐ Ecocardiograma (dejar instrucciones) → realizado por ____________").bold = True

    doc.add_paragraph("\n☐ Envío de muestras (hora: de 11:00 a 13:30)")
    doc.add_paragraph("\nFirmado: _______________________      Fecha: ______________")

    # ==========================================
    # PÁGINA 2: HOJA MÉDICA
    # ==========================================
    doc.add_page_break()
    doc.add_heading(f"HOJA MÉDICA: {datos.get('visita', 'N/A')}", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Protocolo: {datos.get('protocolo', 'ALXN2220-ATTR-CM-301')}     ID: ______________").bold = True
    doc.add_paragraph("☐ Fecha de la visita: ______________\n")

    if proc.get("examen_fisico"):
        doc.add_paragraph("☐ Examen físico completo").bold = True
        doc.add_paragraph("(Ver aspecto general, piel, nariz, orejas, ojos, cuello, garganta, corazón, abdomen, pulmones, sistema vascular, sistema nervioso, sistema musculo esquelético y extremidades)\n")

    if proc.get("nyha"):
        doc.add_paragraph("☐ Clasificación funcional NYHA (seleccionar una):").bold = True
        doc.add_paragraph("   ☐ NYHA I: Asintomático")
        doc.add_paragraph("   ☐ NYHA II: Falta de aire (disnea) a grandes esfuerzos")
        doc.add_paragraph("   ☐ NYHA III: Falta de aire (disnea) a pequeños esfuerzos")
        doc.add_paragraph("   ☐ NYHA IV: Falta de aire (disnea) en reposo (se ahoga estando quieto)\n")

    if proc.get("karnofsky"):
        doc.add_paragraph("☐ Discapacidad (con escala Karnofsky):").bold = True
        doc.add_paragraph("   100% - Actividad normal (capaz de desempeñar actividades), asintomático.")
        doc.add_paragraph("   90% - Actividad normal, con síntomas y signos leves.")
        doc.add_paragraph("   80% - Actividad normal con esfuerzo, síntomas leves.")
        doc.add_paragraph("   70% - Capaz de cuidar de si mismo, pero no realiza trabajo activo.")
        doc.add_paragraph("   60% - En ocasiones necesita ayuda, pero capaz de cuidarse la mayor parte del tiempo.")
        doc.add_paragraph("   50% - Necesita atención médica frecuente y ayuda de otros.")
        doc.add_paragraph("   40% - Con discapacidad, requiere cuidados especiales.")
        doc.add_paragraph("   30% - Discapacidad grave, en condiciones de hospitalización.")
        doc.add_paragraph("   20% - Enfermo grave, necesita tratamiento activo de sostén.")
        doc.add_paragraph("   10% - Paciente decaído o moribundo.")
        doc.add_paragraph("   0% - Paciente fallecido.\n")

    doc.add_paragraph("☐ Medicamentos, terapias o procedimientos simultáneos").bold = True
    doc.add_paragraph("☐ Eventos adversos (AEs)\n").bold = True

    if proc.get("ecg_pre"):
        doc.add_paragraph("☐ Pre-Electrocardiograma de 12 derivaciones").bold = True
        doc.add_paragraph("   ☐ Significancia clínica: _________________________________________")

    if proc.get("ecg_post"):
        doc.add_paragraph("☐ Post-Electrocardiograma de 12 derivaciones").bold = True
        doc.add_paragraph("   ☐ Significancia clínica: _________________________________________")

    if proc.get("infusion"):
        doc.add_paragraph("\n☐ Examen físico breve post infusión:").bold = True
        doc.add_paragraph("Inclusive of general appearance, heart, lungs, skin, musculoskeletal system and extremities and other organs or body systems as clinically indicated should be performed prior to the participant's discharge.")

    doc.add_paragraph("\n\nFirmado (Médico): _______________________      Fecha: ______________")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- 🔍 EXTRACCIÓN DE TEXTO ---
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
    except: texto = "Error en lectura."
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
    st.caption("v5.0 | Motor de Plantillas Clínicas HUJ")

st.markdown('<div class="step-header">📄 Paso 1: Documentación</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">Carga el Protocolo SoE (Las X).</div>', unsafe_allow_html=True)
    f_proto = st.file_uploader("1. SUBIR PROTOCOLO (Apartado SoE)", type=["pdf"])
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="step-header">🔍 Paso 2: Configuración de la Visita</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">Indica la página y el nombre de la visita a generar.</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    p_tabla = c1.text_input("Páginas de la Tabla SoE (ej. 11-16):", "11-16")
    v_proto = c2.text_input("Visita en Protocolo (ej. Visit 17):", "Visit 17")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="step-header">📋 Paso 3: Generación del Checklist</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="step-container"><div class="step-explanation">La IA analizará el protocolo y generará las Hojas de Enfermería y Médica oficiales.</div>', unsafe_allow_html=True)
    
    if st.button("✨ Generar Hoja de Visita Oficial HUJ"):
        if not api_key or not f_proto:
            st.error("⚠️ Faltan documentos o la API Key.")
        else:
            barra_progreso = st.progress(0)
            status_text = st.empty()
            
            try:
                status_text.text("📖 Leyendo la tabla del protocolo...")
                t_soe = extraer_texto_paginas(f_proto.read(), p_tabla)
                barra_progreso.progress(50)
                
                status_text.text("🧠 IA activando los bloques de la plantilla del hospital...")
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(modelo_seleccionado)
                
                prompt = f"""
                Eres el Director Médico del ensayo. Analiza la tabla de protocolo proporcionada para la visita '{v_proto}'.
                
                TABLA SOE: {t_soe}

                Tu única tarea es leer la columna de '{v_proto}' e indicarme qué procedimientos médicos tienen una marca (X, punto, etc).
                Responde ÚNICAMENTE con un JSON en formato booleano (true si está marcado en la tabla para esa visita, false si no).

                Reglas de mapeo para los booleanos:
                - cuestionarios: true si incluye KCCQ, EQ-5D, SF-36, PGIC o PROs.
                - signos_vitales: true si incluye Vital signs o BP/HR.
                - ecg_pre: true si incluye 12-lead ECG.
                - ecg_post: true si el protocolo indica ECG post-infusión.
                - test_6mwt: true si incluye 6-minute walk test o 6MWT.
                - laboratorio: true si incluye clinical laboratory tests, blood, hematology, chemistry, etc.
                - infusion: true si incluye Study intervention infusion o Study drug administration.
                - pk_post: true si incluye PK samples o Pharmacokinetic.
                - eco: true si incluye Echocardiogram.
                - examen_fisico: true si incluye Physical examination (Full o Symptom-directed).
                - nyha: true si incluye NYHA Functional Classification.
                - karnofsky: true si incluye Karnofsky Performance Status score.

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
                    "nyha": true|false,
                    "karnofsky": true|false
                  }}
                }}
                """
                
                res = model.generate_content(prompt)
                doc_word = crear_documento_word_pro(res.text)
                
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
