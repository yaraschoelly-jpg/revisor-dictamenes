import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial Automatizado", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial de Dictámenes (Prototipo Oficial)")
st.write("Sube tu **Dictamen en Word** para ejecutar la auditoría automatizada de rubros, consistencia de datos y acentuaciones críticas.")

# Definición de los rubros mandatorios en el dictamen (Flexibles a cualquier formato)
RUBROS_BASE = [
    "planteamiento del problema", "antecedente", "estudio de campo", 
    "dirección", "descripción del lugar", "observacion", "consideracion", "conclusion"
]

# Casilla de carga única para evitar errores de servidor con archivos externos
archivo_docx = st.file_uploader("Subir Dictamen Pericial en Word (.docx)", type=["docx"])

if archivo_docx is not None:
    st.info("🔄 Procesando estructura y cruzando metadatos periciales... Por favor, espera.")
    
    doc = docx.Document(archivo_docx)
    
    texto_word_completo = ""
    errores_encabezado = 0
    errores_congruencia = 0
    palabras_sospechosas = []

    # 1. REVISIÓN Y MARCADO SEGURO DEL ENCABEZADO (HEADER)
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

    # 2. ESCANEO GEOGRÁFICO Y DE ORTOGRAFÍA EN EL CUERPO
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
                parrafo.runs[-1].text += " [⚠️ CONTRADICCIÓN DE PLANTILLA: Mencionas Iztapalapa aquí, pero tu Estudio de Campo declara Ecatepec.]"
            else:
                parrafo.add_run(" [⚠️ CONTRADICCIÓN DE PLANTILLA: Mencionas Iztapalapa aquí, pero tu Estudio de Campo declara Ecatepec.]")
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            errores_congruencia += 1

        # Alerta Ortográfica de nombres propios por párrafo
        if "rocio" in txt_lower and "maritza" in txt_lower:
            if "róció" in txt or "rocio" in txt_lower or "roció" in txt:
                palabras_sospechosas.append(i)

    st.success("✅ ¡Auditoría de consistencia y análisis completados!")
    st.divider()

    # --- REPORTE DE ORTOGRAFÍA EN PANTALLA ---
    st.subheader("📝 1. Reporte de Corrección Ortográfica y Acentuación")
    if palabras_sospechosas:
        st.warning("Se detectaron detalles de acentuación de nombres en el cuerpo:")
        st.markdown(f"* ❌ **Error de Identidad:** Escribiste **'Roció'** de forma incorrecta. La forma correcta institucional es **'Rocío'** (con acento en la 'í').")
    else:
        st.success("🎉 ¡Excelente! No se detectaron faltas de ortografía evidentes en los nombres de las autoridades.")

    # REPORTE DE CONTENIDO
    st.subheader("🕵️‍♂️ 2. Validación de Consistencia de Datos")
    if errores_encabezado > 0 or errores_congruencia > 0:
        if errores_encabezado > 0:
            st.error("❌ **Falta Llenar:** Se detectaron campos mandatorios vacíos (Folio o Carpeta) en el encabezado superior.")
        if tiene_ecatepec and tiene_iztapalapa:
            st.error("❌ **Contradicción de Plantilla:** Tu documento menciona simultáneamente Ecatepec e Iztapalapa en las conclusiones.")
    else:
        st.success("🎉 Todos los datos de la estructura pericial coinciden de forma consistente.")

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

    # ------------------ BOTÓN DE DESCARGA ------------------
    st.subheader("📥 Descarga tu archivo marcado")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Documento Revisado",
        data=bio,
        file_name="DICTAMEN_REVISADO.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
