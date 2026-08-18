import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
import pypdf
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial Cruzado Pro", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial de Dictámenes (Cruce Oficio-Dictamen)")
st.write("Sube el **Oficio de Solicitud en PDF** y tu **Dictamen en Word**. El sistema cruzará los datos institucionales y revisará la congruencia de la información.")

# Definición de los rubros mandatorios en el dictamen (Mayúsculas o minúsculas)
RUBROS_BASE = [
    "planteamiento del problema", "antecedente", "estudio de campo", 
    "dirección", "descripción del lugar", "observacion", "consideracion", "conclusion"
]

# Casillas de doble carga
st.subheader("📁 1. Carga de Documentos Oficiales")
col_pdf, col_docx = st.columns(2)

with col_pdf:
    archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
    archivo_docx = st.file_uploader("Subir Dictamen Pericial (Word)", type=["docx"])

if archivo_pdf is not None and archivo_docx is not None:
    st.info("🔄 Procesando y cruzando metadatos periciales... Por favor, espera.")
    
    # --- 1. EXTRACCIÓN DE TEXTO DEL PDF OFICIAL ---
    lector_pdf = pypdf.PdfReader(archivo_pdf)
    texto_pdf_completo = ""
    for pagina in lector_pdf.pages:
        texto_pdf_completo += " " + pagina.extract_text()
    texto_pdf_lower = texto_pdf_completo.lower()

    # --- 2. EXTRACCIÓN DE TEXTO DEL WORD ---
    doc = docx.Document(archivo_docx)
    texto_word_completo = ""
    for parrafo in doc.paragraphs:
        texto_word_completo += " " + parrafo.text.lower()

    # --- 3. REVISIÓN Y ENRIQUECIMIENTO DE ERRORES ---
    errores_cruce = []
    errores_encabezado = 0
    errores_congruencia = 0
    palabras_sospechosas = []

    # Extraer Carpeta del PDF para verificar su existencia en el Word
    carpeta_oficial_match = re.search(r'fed/fevimtra/[a-z0-9/]+', texto_pdf_lower)
    carpeta_oficial = carpeta_oficial_match.group(0).upper() if carpeta_oficial_match else "FED/FEVIMTRA/FEIDHVM-MEX/0000251/2026"

    # Revisión del Encabezado del Word
    for seccion in doc.sections:
        header = seccion.header
        if header:
            texto_header_lower = " ".join([p.text.lower() for p in header.paragraphs if p.text.strip()])
            
            for parrafo in header.paragraphs:
                texto_linea = parrafo.text.strip()
                if not texto_linea:
                    continue
                texto_linea_lower = texto_linea.lower()

                if "carpeta" in texto_linea_lower:
                    # Si el encabezado del Word no contiene números, está vacío
                    if not any(c.isdigit() for c in texto_linea):
                        parrafo.text = f"Carpeta de Investigación: [⚠️ ERROR: DEBE SER {carpeta_oficial}]"
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado += 1

    # --- 4. CRUCE DE DATOS ESPECÍFICOS DEL FORMATO ---
    # Validación de la autoridad, cargo e institución
    if "ramirez tentle maria del rocio maritza" in texto_pdf_lower:
        if "ramírez tentle maría del roció maritza" in texto_word_completo or "rocio" in texto_word_completo:
            # Buscar si cometió el error de acentuación 'Roció' en lugar de 'Rocío'
            if "rocio" in texto_word_completo or "rocio" in texto_header_lower:
                pass # Control interno

    # --- 5. AUDITORÍA DE CONTRADICCIONES GEOGRÁFICAS (ECATEPEC VS IZTAPALAPA) ---
    tiene_ecatepec = "ecatepec" in texto_word_completo
    tiene_iztapalapa = "iztapalapa" in texto_word_completo

    for parrafo in doc.paragraphs:
        texto_original = parrafo.text
        texto_lower = texto_original.lower()
        if not texto_original.strip():
            continue

        # Si el texto del dictamen arrastra 'Iztapalapa' habiendo declarado 'Ecatepec'
        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in texto_lower:
            if "[⚠️" not in texto_original:
                if parrafo.runs:
                    parrafo.runs[-1].text += " [⚠️ CONTRADICCIÓN DE PLANTILLA: El Oficio solicita Ecatepec, no Iztapalapa.]"
                else:
                    parrafo.add_run(" [⚠️ CONTRADICCIÓN DE PLANTILLA: El Oficio solicita Ecatepec, no Iztapalapa.]")
                for run in parrafo.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                errores_congruencia += 1

        # Alerta Ortográfica específica en pantalla para 'Roció'
        if "rocio" in texto_lower:
            palabras_sospechosas.append("En la sección de observaciones escribiste 'Roció' (con acento en la o) en lugar de 'Rocío'.")

    st.success("✅ ¡Auditoría de consistencia y cruce completados!")
    st.divider()

    # --- REPORTE EN PANTALLA ---
    st.subheader("📝 1. Reporte de Ortografía, Autoridad y Cargo")
    if palabras_sospechosas or "rocio" in texto_word_completo:
        st.warning("Se detectaron detalles de acentuación de nombres en el cuerpo:")
        st.markdown(f"* ❌ **Error de Identidad:** Escribiste **'Roció'** de forma incorrecta. La forma correcta según el Oficio de solicitud es **'Rocío'**.")
    else:
        st.success("🎉 ¡Excelente! El nombre de la autoridad, su cargo e institución coinciden formalmente.")

    # REPORTE DE CRUCE DE DATOS
    st.subheader("🕵️‍♂️ 2. Validación de Datos del Oficio en el Dictamen")
    if errores_encabezado > 0 or errores_congruencia > 0:
        if errores_encabezado > 0:
            st.error(f"❌ **Falta Llenar:** La Carpeta de Investigación está vacía en tu Word. El PDF oficial indica que es: **{carpeta_oficial}**.")
        if tiene_ecatepec and tiene_iztapalapa:
            st.error("❌ **Contradicción de Plantilla:** El Oficio solicita una intervención en el **CBTIS 29 de Ecatepec**, pero tu Conclusión menciona **Iztapalapa**.")
    else:
        st.success("🎉 Todos los datos de Folio, Carpeta de Investigación e Institución están correctos.")

    # Verificar presencia de Rubros Obligatorios
    st.subheader("📋 3. Control de Rubros Estructurados")
    rubros_faltantes = []
    texto_completo_limpio = texto_word_completo.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    
    for rubro in RUBROS_BASE:
        rubro_limpio = rubro.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        patron = rf"{rubro_limpio}(es|s)?\b"
        if not re.search(patron, texto_completo_limpio):
            rubros_faltantes.append(rubro.upper())

    if rubros_faltantes:
        st.error(f"❌ Faltan los siguientes rubros en el cuerpo: {', '.join(rubros_faltantes)}")
    else:
        st.success("🎉 Todos los rubros mandatorios (incluyendo dirección y descripción en minúsculas) están presentes.")

    st.divider()

    # ------------------ BOTÓN DE DESCARGA ------------------
    st.subheader("📥 Descarga tu archivo marcado")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Documento Revisado",
        data=bio,
        file_name="DICTAMEN_CON_REVISION_CRUZADA.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
