import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial Pro", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial de Dictámenes (Cruce Automatizado)")
st.write("Sube tu **Dictamen en Word**. El sistema lo auditará cruzándolo con las reglas institucionales del Oficio de Solicitud.")

# Definición de los rubros mandatorios en el dictamen (Flexibles a cualquier formato)
RUBROS_BASE = [
    "planteamiento del problema", "antecedente", "estudio de campo", 
    "dirección", "descripción del lugar", "observacion", "consideracion", "conclusion"
]

# Casilla de carga única de Word (A prueba de errores de servidor)
archivo_docx = st.file_uploader("Subir Dictamen Pericial en Word (.docx)", type=["docx"])

if archivo_docx is not None:
    st.info(" spy🔍 Ejecutando cruce de datos institucionales y formato... Por favor, espera.")
    
    doc = docx.Document(archivo_docx)
    
    texto_word_completo = ""
    errores_encabezado = 0
    errores_congruencia = 0
    errores_institucionales = []
    error_rocio_párrafos = []

    # DATOS OFICIALES DE REFERENCIA ESTRICTA DEL OFICIO
    FOLIO_OFICIAL = "fgr-aic-pfm-uinp-diedcs-sa-017608-2026"
    CARPETA_OFICIAL = "FED/FEVIMTRA/FEIDHVM-MEX/0000251/2026"
    AUTORIDAD_OFICIAL = "ramírez tentle maría del rocío maritza"

    # 1. REVISIÓN Y MARCADO SEGURO DEL ENCABEZADO (HEADER)
    for seccion in doc.sections:
        header = seccion.header
        if header:
            texto_header_lower = " ".join([p.text.lower() for p in header.paragraphs if p.text.strip()])

            for parrafo in header.paragraphs:
                texto_linea = parrafo.text.strip()
                if not texto_linea:
                    continue
                texto_linea_lower = texto_linea.lower()

                # Ignorar renglones fijos institucionales del encabezado
                if any(t in texto_linea_lower for t in ["agencia", "centro federal", "unidad de", "especialidad"]):
                    continue

                # Validar Folio Cruzado
                if "folio" in texto_linea_lower:
                    if FOLIO_OFICIAL not in texto_linea_lower:
                        parrafo.text = f"Número de folio: [⚠️ ERROR: DEBE SER FGR-AIC-PFM-UINP-DIEDCS-SA-017608-2026]"
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado += 1

                # Validar Carpeta Cruzada
                elif "carpeta" in texto_linea_lower or "investigación" in texto_linea_lower:
                    if "0000251" not in texto_linea_lower:
                        parrafo.text = f"Carpeta de Investigación: [⚠️ ERROR: DEBE SER {CARPETA_OFICIAL}]"
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado += 1

    # 2. ESCANEO DEL CUERPO DEL DOCUMENTO
    tiene_ecatepec = False
    tiene_iztapalapa = False

    for i, parrafo in enumerate(doc.paragraphs, start=1):
        txt = parrafo.text.strip()
        if not txt:
            continue
        txt_lower = txt.lower()
        texto_word_completo += " " + txt_lower
        
        if "ecatepec" in txt_lower:
            tiene_ecatepec = True
        if "iztapalapa" in txt_lower:
            tiene_iztapalapa = True

        # Marcar Contradicción Geográfica Directa (Ecatepec vs Iztapalapa)
        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in txt_lower and "[⚠️" not in txt:
            if parrafo.runs:
                parrafo.runs[-1].text += " [⚠️ CONTRADICCIÓN DE PLANTILLA: Tu Estudio de Campo declara Ecatepec, no Iztapalapa.]"
            else:
                parrafo.add_run(" [⚠️ CONTRADICCIÓN DE PLANTILLA: Tu Estudio de Campo declara Ecatepec, no Iztapalapa.]")
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            errores_congruencia += 1

        # Alerta Ortográfica estricta de acentuación 'Roció' vs 'Rocío'
        if "rocio" in txt_lower and "maritza" in txt_lower:
            if "roció" in txt_lower or "rocio" in txt_lower:
                error_rocio_párrafos.append(i)

    # 3. VERIFICAR CRUCE DE IDENTIDADES INSTITUCIONALES EN EL TEXTO COMPLETADO
    if "maría del rocío maritza" not in texto_word_completo and "ramírez tentle" not in texto_word_completo:
        errores_institucionales.append("❌ **Nombre de Autoridad Omitido:** Falta escribir el nombre completo de la Agente (*Ramírez Tentle María del Rocío Maritza*).")
    if "oficial investigador" not in texto_word_completo:
        errores_institucionales.append("❌ **Cargo Omitido o Discrepante:** Falta el cargo exacto (*Oficial Investigador B*).")
    if "policía federal ministerial" not in texto_word_completo:
        errores_institucionales.append("❌ **Institución Discrepante:** Falta hacer mención a la *Policía Federal Ministerial*.")

    st.success("✅ ¡Auditoría de consistencia y cruce de datos completados!")
    st.divider()

    # --- REPORTE DE ORTOGRAFÍA EN PANTALLA ---
    st.subheader("📝 1. Reporte de Corrección Ortográfica y Acentuación")
    if error_rocio_párrafos:
        st.warning("Se detectaron detalles de acentuación críticos de nombres propios:")
        st.markdown(f"* En tus párrafos de identificación escribiste **'Roció'** de forma incorrecta. La forma oficial es **'Rocío'** (con acento en la 'í').")
    else:
        st.success("🎉 ¡Excelente! No se detectaron faltas de ortografía evidentes en los nombres del personal.")

    # REPORTE DE CRUCE DE DATOS DEL OFICIO
    st.subheader("🕵️‍♂️ 2. Validación Cruzada con el Oficio de Solicitud")
    if errores_encabezado > 0 or errores_congruencia > 0 or errores_institucionales:
        if errores_encabezado > 0:
            st.error("❌ **Error en Encabezado:** Tu folio o carpeta están vacíos o no coinciden con los datos oficiales del Oficio.")
        if tiene_ecatepec and tiene_iztapalapa:
            st.error("❌ **Contradicción de Plantilla Geográfica:** El texto menciona simultáneamente Ecatepec e Iztapalapa (Revisar la Conclusión).")
        for err_inst in errores_institucionales:
            st.error(err_inst)
    else:
        st.success("🎉 ¡Espectacular! Todos los datos de Folio, Carpeta, Autoridad, Cargo e Institución coinciden formalmente con el Oficio oficial.")

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
        st.error(f"❌ Faltan los siguientes rubros obligatorios en el cuerpo del Word: {', '.join(rubros_faltantes)}")
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
        file_name="DICTAMEN_REVISADO_COMPLETO.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
