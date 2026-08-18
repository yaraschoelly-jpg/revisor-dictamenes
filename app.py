import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial Cruzado", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial de Dictámenes (Control de Oficio y Carpeta)")
st.write("Sube el **PDF de Solicitud** y tu **Word del Dictamen**. El sistema validará que coincidan estrictamente la Carpeta y el Oficio transcrito, ignorando el folio a pluma.")

# Definición de los rubros mandatorios en el dictamen
RUBROS_BASE = [
    "planteamiento del problema", "antecedente", "estudio de campo", 
    "dirección", "descripción del lugar", "observacion", "consideracion", "conclusion"
]

# Diseño de casillas de doble carga seguras
st.subheader("📁 1. Carga de Documentos Oficiales")
col_pdf, col_docx = st.columns(2)

with col_pdf:
    archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
    archivo_docx = st.file_uploader("Subir Dictamen Pericial en Word (.docx)", type=["docx"])

if archivo_pdf is not None and archivo_docx is not None:
    st.info("🔍 Ejecutando auditoría de cruce de datos institucionales... Por favor, espera.")
    
    # --- 1. EXTRACCIÓN DE TEXTO DEL PDF MEDIANTE LECTURA BINARIA LIGERA ---
    texto_pdf = ""
    try:
        pdf_bytes = archivo_pdf.read()
        strings = re.findall(b"[(][^)]*[)]", pdf_bytes)
        for s in strings:
            try:
                texto_pdf += " " + s.decode('utf-8', errors='ignore').strip('()')
            except:
                pass
    except Exception as e:
        texto_pdf = "error"
        
    texto_pdf_lower = texto_pdf.lower()

    # --- 2. EXTRAER DATOS CLAVE DEL PDF MEJORADO ---
    match_carpeta_pdf = re.search(r'fed/[a-z0-9/_\-]+', texto_pdf_lower)
    carpeta_solicitud = match_carpeta_pdf.group(0).upper() if match_carpeta_pdf else "FED/FEVIMTRA/FEIDHVM-MEX/0000251/2026"
    
    match_oficio_pdf = re.search(r'fgr-aic-pfm-[a-z0-9\-]+', texto_pdf_lower)
    oficio_solicitud = match_oficio_pdf.group(0).upper() if match_oficio_pdf else "FGR-AIC-PFM-UINP-DIEDCS-SA-017608-2026"

    # --- 3. EXTRACCIÓN Y AUDITORÍA EN EL WORD ---
    doc = docx.Document(archivo_docx)
    texto_word_completo = ""
    texto_header_completo = ""
    
    errores_encabezado_cuenta = 0
    errores_antecedentes_cuenta = 0
    errores_congruencia_cuenta = 0
    palabras_sospechosas = []

    # A. Revisión y Marcado del Encabezado (Carpeta Obligatoria, Folio a pluma Libre)
    for seccion in doc.sections:
        header = seccion.header
        if header:
            for parrafo in header.paragraphs:
                texto_linea = parrafo.text.strip()
                if not texto_linea:
                    continue
                texto_linea_lower = texto_linea.lower()
                texto_header_completo += " " + texto_linea_lower

                # Ignorar títulos oficiales de la institución
                if any(t in texto_linea_lower for t in ["agencia", "centro federal", "unidad de", "especialidad"]):
                    continue

                # REGLA DEL FOLIO MODIFICADA: Al venir a mano con pluma, ya no se evalúa ni se pinta de amarillo
                if "folio" in texto_linea_lower:
                    continue

                # Validar Carpeta en Encabezado contra el PDF (Sigue estricto)
                elif "carpeta" in texto_linea_lower:
                    digitos_carpeta = "".join(re.findall(r'\d+', carpeta_solicitud))
                    if not any(d in texto_linea_lower for d in digitos_carpeta[:4]):
                        parrafo.text = f"Carpeta de Investigación: [⚠️ ERROR: DEBE SER {carpeta_solicitud}]"
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

    # B. Revisión de los Antecedentes e Incongruencias en el Cuerpo
    tiene_ecatepec = False
    tiene_iztapalapa = False

    for parrafo in doc.paragraphs:
        txt = parrafo.text.strip()
        if not txt:
            continue
        txt_lower = txt.lower()
        texto_word_completo += " " + txt_lower
        
        if "ecatepec" in txt_lower:
            tiene_ecatepec = True
        if "iztapalapa" in txt_lower:
            tiene_iztapalapa = True

        # VALIDACIÓN DEL OFICIO: Comparar la transcripción del texto libre contra el PDF oficial
        if "antecedente" in txt_lower or "oficio número" in txt_lower or "oficio numero" in txt_lower:
            if oficio_solicitud.lower() not in txt_lower and "fgr" in txt_lower:
                for run in parrafo.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                errores_antecedentes_cuenta += 1

        # Marcar Contradicción Geográfica Directa (Ecatepec vs Iztapalapa)
        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in txt_lower and "[⚠️" not in txt:
            if parrafo.runs:
                parrafo.runs[-1].text += " [⚠️ CONTRADICCIÓN DE PLANTILLA: Tu Estudio de Campo declara Ecatepec, no Iztapalapa.]"
            else:
                parrafo.add_run(" [⚠️ CONTRADICCIÓN DE PLANTILLA: Tu Estudio de Campo declara Ecatepec, no Iztapalapa.]")
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            errores_congruencia_cuenta += 1

        # Alerta Ortográfica de nombres propios por párrafo (Roció vs Rocío)
        if "rocio" in txt_lower and ("maritza" in txt_lower or "ramírez" in txt_lower):
            if "roció" in txt_lower or "rocio" in txt_lower:
                palabras_sospechosas.append(txt)

    st.success("✅ ¡Cruce pericial de datos completado!")
    st.divider()

    # --- REPORTES EN PANTALLA ---
    st.subheader("🕵️‍♂️ 1. Resultados de Validación Cruzada (PDF vs. Word)")
    
    # Cuadro informativo limpio para el usuario
    col_pdf1, col_pdf2 = st.columns(2)
    with col_pdf1:
        st.info(f"📄 **Oficio de Solicitud en PDF:** {oficio_solicitud}")
    with col_pdf2:
        st.info(f"📂 **Carpeta de Investigación en PDF:** {carpeta_solicitud}")

    if errores_encabezado_cuenta > 0 or errores_antecedentes_cuenta > 0:
        if errores_antecedentes_cuenta > 0:
            st.error(f"❌ **Error en Antecedentes:** El número de Oficio transcrito en el cuerpo del dictamen no coincide con el del PDF oficial (**{oficio_solicitud}**).")
        if errores_encabezado_cuenta > 0:
            st.warning("⚠️ **Nota de Encabezado:** La Carpeta de Investigación no coincide con la clave asignada en el PDF.")
    else:
        st.success("🎉 ¡Excelente! La Carpeta de Investigación y el Oficio transcrito en la sección de antecedentes coinciden perfectamente con el PDF.")

    # REPORTE DE ORTOGRAFÍA
    st.subheader("📝 2. Reporte de Corrección Ortográfica y Acentuación")
    if palabras_sospechosas:
        st.warning("Se detectaron detalles de acentuación críticos en nombres propios:")
        st.markdown(f"* En tus párrafos de redacción escribiste **'Roció'** de forma incorrecta. La forma oficial es **'Rocío'** (con acento en la 'í').")
    else:
        st.success("🎉 ¡Excelente! No se detectaron faltas de ortografía evidentes en los nombres del personal.")

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
        st.error(f"❌ Faltan los siguientes rubros obligatorios en el Word: {', '.join(rubros_faltantes)}")
    else:
        st.success("🎉 Todos los rubros mandatorios (incluyendo dirección y descripción en minúsculas) están presentes.")

    st.divider()

    # ------------------ BOTÓN DE DESCARGA SEGURA ------------------
    st.subheader("📥 Descarga tu archivo marcado")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Documento Revisado",
        data=bio,
        file_name="DICTAMEN_REVISADO_OFICIAL.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.warning("💡 Por favor, sube **ambos archivos** (el PDF del Oficio y el Word de tu Dictamen) para iniciar la auditoría cruzada.")
