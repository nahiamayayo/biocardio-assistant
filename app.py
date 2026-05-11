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
    /* Fondo principal y fuentes */
    .stApp {
        background-color: #f4f7f4;
    }
    /* Estilo de los contenedores de pasos (Tarjetas) */
    .step-container {
        background-color: white;
        padding: 25px 30px;
        border-radius: 15px;
        border-left: 8px solid #007d32;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        margin-bottom: 35px;
    }
    /* Títulos FUERA de la tarjeta */
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
    /* Explicación DENTRO de la tarjeta */
    .step-explanation {
        color: #555;
        font-size: 16px;
        margin-bottom: 20px;
        font-style: italic;
        border-bottom: 1px solid #e0e0e0;
        padding-bottom: 10px;
    }
    /* Botón principal estilizado */
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
    .stButton>button:hover {
        background-color: #005a24 !important;
        box-shadow: 0 4px 15px rgba(0,125,50,0.3);
    }
    </style>
""", unsafe_allow_html=True)

# --- 📄 GENERADOR DEL WORD (INCLUYE ECG Y 6MWT) ---
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
    style.font.color.rgb = RGBColor(0, 0, 0) 
    
    for i in range(1, 4):
        if f'Heading {i}' in doc.styles:
            h_style = doc.styles[f'Heading {i}']
            h_style.font.name = 'Calibri'
            h_style.font.color.rgb = RGBColor(0, 0, 0)

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

    if procedimientos:
        for titulo_cat, id_cat in categorias_ordenadas:
            items = [p for p in procedimientos if p.get("categoria") == id_cat]
            if items:
                doc.add_heading(titulo_cat, level=1)
                for item in items:
                    p = doc.add_paragraph()
                    p.add_run(f"[  ] {item.get('procedimiento')} ").bold = True
                    
                    proc_nombre = item.get('procedimiento', '').lower()

                    # --- INYECCIÓN ESPECIAL: ECG ---
                    if "ecg" in proc_nombre or "electrocardiograma" in proc_nombre:
                        ecg_p = doc.add_paragraph()
                        ecg_p.paragraph_format.left_indent = Pt(20)
                        ecg_p.add_run("• Posición supino\n• Frecuencia cardíaca: ..........\n• PR: ..........\n• QRS: ..........\n• QT: ..........\n• QTc: ..........")

                    # --- INYECCIÓN ESPECIAL: TEST DE 6 MINUTOS ---
                    if "6 min" in proc_nombre or "6-min" in proc_nombre:
                        test_p = doc.add_paragraph()
                        test_p.paragraph_format.left_indent = Pt(20)
                        test_p.add_run("(después de sangre y cuestionarios, pero antes de la infusión)\n").italic = True
                        test_p.add_run("• Coger material y hoja, escribir en inglés en la hoja → realizado por: ________________\n")
                        test_p.add_run("• Rellenar en Course Name: 6MWT 2F; y en Course Lenght: 20m\n")
                        test_p.add_run("• Todas las casillas deben estar completas. Si no hiciste alguna casilla poner: N/D.")

                    if item.get("detalles"):
                        det_p = doc.add_paragraph(item.get("detalles"))
                        det_p.paragraph_format.left_indent = Pt(20)

                    tiempos = item.get("tiempos_especificos", [])
                    if "signos vitales" in proc_nombre or "vital signs" in proc_nombre:
                        if not tiempos:
                            tiempos = ["________________", "________________"]
                        table = doc.add_table(rows=len(tiempos)+1, cols=8)
                        table.style = 'Table Grid'
                        headers = ["Tiempo", "Hora", "PA Sist.", "PA Diast.", "Pulso", "Resp", "O2 (%)", "Temp."]
                        for i, head in enumerate(headers):
                            table.cell(0, i).text = head
                            table.cell(0, i).paragraphs[0].runs[0].bold = True
                        for i, t_val in enumerate(tiempos):
                            table.cell(i+1, 0).text = t_val

                    if id_cat == "laboratorio":
                        chk = doc.add_paragraph("   [  ] Muestra extraída y procesada según el manual de laboratorio\n   [  ] Muestra enviada o almacenada correctamente al laboratorio central")
                        chk.paragraph_format.left_indent = Pt(20)

    doc.add_paragraph("\n\n")
    firma_table = doc.add_table(rows=1, cols=2)
    firma_table.cell(0, 0).text = "Firma del Investigador / Enfermería:\n\n___________________________"
    firma_table.cell(0, 1).text = "Fecha:\n\n___________________________"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- 🔍 EXTRACCIÓN DE TEXTO ---
def extraer_texto_paginas(pdf_bytes, paginas_str=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    paginas_a_leer = []
    if paginas_str:
        p_set = set()
        for b in paginas_str.split(','):
            if '-' in b:
                s, e = b.split('-')
                p_set.update(range(int(s), int(e)+1))
            else: p_set.add(int(b))
        paginas_a_leer = sorted(list(p_set))
    else:
        paginas_a_leer = range(1, len(doc) + 1)

    for p in paginas_a_leer:
        if 0 < p <= len(doc):
            page = doc[p-1]
            try: texto += f"\n--- PÁGINA {p} ---\n" + page.get_text("markdown") + "\n"
            except: texto += f"\n--- PÁGINA {p} ---\n" + page.get_text("text") + "\n"
    return texto

# --- 💉 TEXTO INFUSIÓN ---
texto_infusion_obligatorio = """*La infusión debe durar 1 hora.
*No diluir el fármaco.
*Limpiar las vías ANTES y DESPUÉS de la infusión con dextrosa 5%: Limpiar la vía con dextrosa 5% antes de la infusión solamente si se le ha infundido algún otro fluido o fármaco al paciente. Si la línea de infusión es nueva, no hace falta hacer este lavado previo a la infusión con dextrosa 5%. 
*Cuando se recoja la medicación de farmacia, se debe completar el IMP Transfer Log, rellenado por farmacia y entregado en el momento de la dispensación, firmado por el equipo unblinded y blinded.
*Administración del fármaco:
1. Revisar visualmente la bolsa de infusión que esté enmascarada. NO SACAR la bolsa de cegado. 
2. Revisar la caducidad de la bolsa de infusión preparada. Si está caducada, no administrar.
3. Conectar el sistema de tubos de infusión a la bolsa. Recordad que hay administrar el fármaco con un filtro de línea de 0.2 micras.
4. Conectar el sistema de tubos a la bomba de infusión y seguir la velocidad determinada por la tabla de abajo.
5. La hora de inicio de la infusión es cuando la bomba ha empezado a infundir. Registrar hora y el brazo. Infundir toda la bolsa. Durante la infusión recoger los signos vitales cada 15 minutos (+- 5 minutos), presión arterial sistólica, diastólica, pulsaciones, respiratory rate, saturación de oxigeno y temperatura.
6. Limpiar la vía con dextrosa 5% (5 minutos) con la misma velocidad de la medicación.
7. Quitar la bolsa de infusión. Registrar la hora (incluyendo el flush con dextrosa 5%) en la hoja de administración de la medicación.

☐ Infusión del fármaco → Dosis según el peso del paciente
□ HORA DE INICIO DE LA INFUSIÓN: ____________
□ HORA DE FIN DE LA INFUSIÓN: ____________
□ Rellenar la hoja de administración de la medicación (IMP transfer log). Revisar la bolsa de cegado.
□ Lavar con 5% dextrosa al final de la infusión."""

# --- 🖥️ INTERFAZ WEB STREAMLIT ---

# Cabecera
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try:
        st.image("huj.png", width=140)
    except:
        st.markdown("## 🏥")

with col_titulo:
    st.markdown("<h1 style='color: #007d32; margin-top: 10px;'>BioCardio Clinical Assistant</h1>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 18px; color: #555;'>Herramienta profesional de gestión de ensayos clínicos</p>", unsafe_allow_html=True)

st.divider()

# Sidebar
with st.sidebar:
    st.markdown("### 🔑 Acceso")
    st.markdown("""
    Para utilizar la IA, necesitas una **API Key personal**. 
    Puedes generarla gratuitamente en:
    [Google AI Studio ↗](https://aistudio.google.com/app/apikey)
    """)
    api_key = st.text_input("Introduce tu API Key:", type="password")
    
    st.divider()
    modelo_seleccionado = st.selectbox("Modelo de Inteligencia Artificial:", ["gemini-3.1-flash-lite-preview"], help="Modelo optimizado para protocolos médicos.")
    
    st.divider()
    st.caption("v3.3 | Hospital Universitario de Jaén")

# Paso 1
st.markdown('<div class="step-header">📄 Paso 1: Documentación</div>', unsafe_allow_html=True)
with st.container():
    st.markdown("""
    <div class="step-container">
        <div class="step-explanation">Carga el Protocolo SoE y los manuales de laboratorio para empezar.</div>
    """, unsafe_allow_html=True)
    
    f_proto = st.file_uploader("Protocolo del Ensayo (SoE)", type=["pdf"])
    f_labs = st.file_uploader("Manuales o Logs de Laboratorio (Puedes adjuntar varios)", type=["pdf"], accept_multiple_files=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

# Paso 2
st.markdown('<div class="step-header">🔍 Paso 2: Configuración de la Visita</div>', unsafe_allow_html=True)
with st.container():
    st.markdown("""
    <div class="step-container">
        <div class="step-explanation">Indica las páginas de la tabla de procedimientos (SoE) y los detalles técnicos (Study Assessments).</div>
    """, unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        p_tabla = st.text_input("Páginas de la Tabla SoE (ej. 11-16):", "11-16")
        v_proto = st.text_input("Visita en Protocolo (ej. Visit 10):", "Visit 10")
    with c2:
        # CAMBIO AQUÍ: Glosario por Study Assessments
        p_assessments = st.text_input("Páginas Study Assessments (ej. 40-60):", "40-60")
        v_lab = st.text_input("Visita en Manual Lab (ej. Visit 10):", "Visit 10")
        
    st.markdown("</div>", unsafe_allow_html=True)

# Paso 3
st.markdown('<div class="step-header">📋 Paso 3: Generación del Checklist</div>', unsafe_allow_html=True)
with st.container():
    st.markdown("""
    <div class="step-container">
        <div class="step-explanation">Procesa la información y genera el documento oficial de enfermería/médico.</div>
    """, unsafe_allow_html=True)
    
    if st.button("Generar Hoja de Visita Profesional"):
        if not api_key:
            st.error("⚠️ Introduce la API Key en el menú lateral.")
        elif not f_proto:
            st.error("⚠️ Adjunta al menos el protocolo (SoE).")
        else:
            barra_progreso = st.progress(0)
            status_text = st.empty()
            
            try:
                status_text.text("📖 Leyendo tablas y procedimientos del protocolo...")
                p_bytes = f_proto.read()
                texto_tabla_proto = extraer_texto_paginas(p_bytes, p_tabla)
                # CAMBIO AQUÍ: Extraemos el texto de Assessments
                texto_assessments = extraer_texto_paginas(p_bytes, p_assessments)
                barra_progreso.progress(30)
                
                status_text.text("🧪 Analizando documentos de laboratorio...")
                texto_laboratorios = ""
                if f_labs:
                    for f in f_labs:
                        texto_laboratorios += f"\n--- DOC: {f.name} ---\n" + extraer_texto_paginas(f.read(), "")
                barra_progreso.progress(60)
                
                status_text.text("🧠 IA maquetando la hoja clínica...")
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(modelo_seleccionado)
                
                prompt = f"""
                Eres un Data Manager Clínico. Tu tarea es LEER las tablas y redactar un JSON para la visita '{v_proto}' (Visita Lab: '{v_lab}').

                --- TEXTO DE LA TABLA DEL PROTOCOLO (SoE) ---
                {texto_tabla_proto}

                --- TEXTO DE DETALLES TÉCNICOS (STUDY ASSESSMENTS) ---
                {texto_assessments}

                --- TEXTO DE LOS DOCUMENTOS DE LABORATORIO ---
                {texto_laboratorios[:100000]}

                INSTRUCCIONES:
                1. TABLA DEL PROTOCOLO: Busca la visita '{v_proto}'. Lista todos los procedimientos marcados con (X).
                2. DETALLES TÉCNICOS: Usa el texto de 'STUDY ASSESSMENTS' para añadir instrucciones precisas en el campo 'detalles' de cada procedimiento (ej. si el ECG es en supino, si los signos vitales requieren reposo previo, etc.).
                3. LABORATORIO (NO AGRUPAR): Busca '{v_lab}' en el texto de lab. Crea UN BLOQUE INDIVIDUAL para cada analito.
                   Detalles: "- Panel: [Nombre]\\n- Tubo: [Volumen] tapón [COLOR], [procesado]".
                4. INFUSIÓN: Si se requiere "Infusion" para esta visita, añade en 'administracion' el bloque de texto exacto:
                   {texto_infusion_obligatorio}
                   Y en signos vitales los tiempos de 15, 30, 45 y 60 min.
                5. CUESTIONARIOS: Incluir Peso, KCCQ, EQ-5D, SF-36, PGIC en 'pre_dosis' si tocan.

                JSON:
                {{ "visita": "{v_proto}", "procedimientos_finales": [ {{"procedimiento": "...", "categoria": "pre_dosis|laboratorio|administracion|post_dosis|continuos|general", "detalles": "...", "tiempos_especificos": []}} ] }}
                """
                
                res = model.generate_content(prompt)
                doc_word = crear_documento_word_pro(res.text)
                
                barra_progreso.progress(100)
                status_text.success("✅ Hoja de visita generada con éxito.")
                
                st.download_button(
                    label="⬇️ Descargar Documento Word",
                    data=doc_word,
                    file_name=f"Checklist_{v_proto.replace(' ','_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error: {e}")
                
    st.markdown("</div>", unsafe_allow_html=True)
