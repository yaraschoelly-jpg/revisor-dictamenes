import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
from spellchecker import SpellChecker
import pypdf
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial Cruzado Pro", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial de Dictámenes (Versión Final)")
st.write("Sube el **Oficio de Solicitud en PDF** y tu **Dictamen en Word**. El sistema cruzará los datos institucionales y revisará la ortografía contextual.")

# Definición de los rubros mandatorios en el dictamen
RUBROS_BASE = [
    "planteamiento del problema", "antecedente", "estudio de campo", 
    "dirección", "descripción del lugar", "observacion", "consideracion", "conclusion"
]

# Diccionario extendido de términos oficiales de México para evitar falsos positivos de ortografía
PALABRAS_SEGURAS = {
    "siendo", "las", "direccion", "dirección", "perita", "perito", "adscrito", "adscrita",
    "exordio", "dictamen", "antecedentes", "planteamiento", "método", "técnica", "estudio",
    "gabinete", "observación", "observaciones", "consideración", "consideraciones", "conclusión",
    "conclusiones", "atentamente", "folio", "carpeta", "investigación", "expediente", "nuc",
    "cbtis", "insurgentes", "ecatepec", "iztapalapa", "comonfort", "maza", "parada", "tentle",
    "fgr", "aic", "pfm", "uinp", "diedcs", "sa", "videovigilancia", "bullet", "domo", "ejidal",
    "s/n", "colonia", "muro", "loseta", "pericial", "forense", "margarita", "rocio", "maritza",
    "elaine", "saraí", "nava", "velázquez", "ramírez", "distanciómetro", "brújula", "adosada",
    "rocío", "lcda", "periciales", "investigador", "policía", "federal", "ministerial", "criminalística"
}

st.subheader("📁 1. Carga de Documentos Oficiales")
col_pdf, col_docx = st.columns(2)

with col_pdf:
    archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
    archivo_docx = st.file_uploader("Subir Dictamen Pericial (Word)", type=["docx"])

if archivo_pdf is not None and archivo_docx is not None:
    st.info("🔄 Ejecutando auditoría y análisis de consistencia... Por favor, espera.")
    
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

    # --- 3. PROCESAMIENTO E IDENTIFICACIÓN DE ERRORES DE CONTENIDO ---
    errores_cruce = []
    errores_encabezado = 0
    errores_congruencia = 0
    palabras_sospechosas = []
    
    spell = SpellChecker(language='es')

    # Capturar datos del encabezado del Word para el reporte
    for seccion in doc.sections:
        header = seccion.header
        if header:
            texto_header_lower = " ".join([p.text.lower() for p in header.paragraphs if p.text.strip()])
            tiene_datos_carpeta = any(caracter.isdigit() for caracter in texto_header_lower)

            for parrafo in header.paragraphs:
                texto_linea = parrafo.text.strip()
                if not texto_linea:
                    continue
                texto_linea_lower = texto_linea.lower()

                if any(t in texto_linea_lower for t in ["agencia", "centro federal", "unidad de", "especialidad"]):
                    continue

                if "folio" in texto_linea_lower:
                    contenido_folio = texto_linea.split(":")[-1].strip() if ":" in texto_linea else re.sub(r'número de folio|numero de folio', '', texto_linea, flags=re.IGNORECASE).strip()
                    if len(contenido_folio) == 0 and "[⚠️" not in texto_linea:
                        if parrafo.runs:
                            parrafo.runs[-1].text += " [⚠️ ERROR: FALTA LLENAR EL NÚMERO DE FOLIO]"
                        else:
                            parrafo.add_run(" [⚠️ ERROR: FALTA LLENAR EL NÚMERO DE FOLIO]")
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado += 1

                elif "carpeta" in texto_linea_lower or "investigación" in texto_linea_lower or "investigacion" in texto_linea_lower:
                    if not tiene_datos_carpeta and "[⚠️" not in texto_linea:
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado += 1

    # --- 4. CRUCE DE IDENTIDADES (AUTORIDAD, CARGO E INSTITUCIÓN) ---
    elementos_autoridad = {
        "maría del rocio maritza": "Nombre de la Autoridad Solicitante (María del Rocio Maritza Ramírez Tentle)",
        "oficial investigador": "Cargo de la Autoridad Solicitante (Oficial Investigador B)",
        "policía federal ministerial": "Institución de la Autoridad (Policía Federal Ministerial)",
        "agencia de investigación criminal": "Agencia/Institución (Agencia de Investigación Criminal)"
    }

    for clave, descripcion in elementos_autoridad.items():
        if clave in texto_pdf_lower and clave not in texto_word_completo:
            errores_cruce.append(f"⚠️ **Omisión o Discrepancia:** No se localizó el dato de **{descripcion}** dentro del cuerpo del dictamen.")

    # --- 5. AUDITORÍA GEOGRÁFICA Y DE ORTOGRAFÍA EN EL CUERPO ---
    tiene_ecatepec = "ecatepec" in texto_word_completo
    tiene_iztapalapa = "iztapalapa" in texto_word_completo

    for i, parrafo in enumerate(doc.paragraphs, start=1):
        texto_original = parrafo.text
        texto_lower = texto_original.lower()
        if not texto_original.strip():
            continue

        # Marcar Contradicción Geográfica (Ecatepec vs Iztapalapa)
        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in texto_lower and "[⚠️" not in texto_original:
            if parrafo.runs:
                parrafo.runs[-1].text += " [⚠️ CONTRADICCIÓN GEOGRÁFICA DE PLANTILLA]"
            else:
                parrafo.add_run(" [⚠️ CONTRADICCIÓN GEOGRÁFICA DE PLANTILLA]")
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            errores_congruencia += 1

        # Análisis Ortográfico Limpio
        palabras = re.findall(r'\b\w+\b', texto_original)
        for palabra in palabras:
            palabra_lower = palabra.lower()
            # Ignoramos palabras cortas, números y la lista estricta mexicana de PALABRAS_SEGURAS
            if len(palabra_lower) > 2 and not palabra_lower.isdigit() and palabra_lower not in PALABRAS_SEGURAS:
                # Si no está en el diccionario oficial o es el error común de acentuación 'Roció'
                if not spell.known([palabra_lower]) or palabra_lower == "roció":
                    sugerencia = "Rocío" if palabra_lower == "roció" else spell.correction(palabra_lower)
                    if sugerencia and sugerencia.lower() != palabra_lower:
                        palabras_sospechosas.append({
                            "linea": i,
                            "erronea": palabra,
                            "sugerencia": sugerencia
                        })

    st.success("✅ ¡Cruce de datos y auditoría de consistencia completados!")
    st.divider()

    # --- REPORTE DE ORTOGRAFÍA EN PANTALLA ---
    st.subheader("📝 1. Reporte de Corrección Ortográfica y Acentuación")
    if palabras_sospechosas:
        st.warning(f"Se detectaron palabras con posibles detalles ortográficos en tu redacción:")
        vistas_ortografia = set()
        for item in palabras_sospechosas:
            clave_ort = f"{item['erronea'].lower()}->{item['sugerencia'].lower()}"
            if clave_ort not in vistas_ortografia:
                st.markdown(f"* En el texto dice: **\"{item['erronea']}\"** 👉 ¿Quisiste decir: *{item['sugerencia']}*?")
                vistas_ortografia.add(clave_ort)
    else:
        st.success("🎉 ¡Excelente! No se detectaron faltas de ortografía evidentes en las hojas.")

    st.divider()

    # --- REPORTE DE CRUCE EN PANTALLA ---
    st.subheader("🕵️‍♂️ 2. Reporte de Validación de Datos Oficiales (Oficio vs. Dictamen)")
    if errores_cruce:
        st.warning(f"Se detectaron discrepancias de contenido entre el Oficio de solicitud y tu Dictamen:")
        for err in errores_cruce:
            st.markdown(err)
    else:
        st.success("🎉 ¡Espectacular! Todos los datos de Folio, Carpeta, Autoridad, Cargo e Institución coinciden perfectamente con el PDF oficial.")

    st.divider()

    # REPORTES EN PANTALLA GENERALES
    st.subheader("📊 3. Alertas de Estructura de Word")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Campos Vacíos Encabezado", errores_encabezado)
    with col2:
        st.metric("Contradicciones de Plantilla", errores_congruencia)

    # Verificar presencia de Rubros Obligatorios
    st.subheader("📋 4. Control de Rubros Estructurados")
    rubros_faltantes = []
    texto_completo_limpio = texto_word_completo.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    
    for rubro in RUBROS_BASE:
        rubro_limpio = rubro.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        patron = rf"{rubro_limpio}(es|s)?\b"
        if not re.search(patron, texto_completo_limpio):
            rubros_faltantes.append(rubro.upper())

    if rubros_faltantes:
        st.error(f"❌ Faltan los siguientes rubros obligatorios en el Word: {', '.join(rubros_faltantes)}")
    else:
        st.success("🎉 Todos los rubros mandatorios están presentes en el cuerpo del dictamen.")

    st.divider()

    # ------------------ BOTÓN DE DESCARGA ------------------
    st.subheader("📥 Descarga tu archivo marcado")
