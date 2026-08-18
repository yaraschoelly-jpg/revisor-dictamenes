import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial de Dictámenes", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial de Dictámenes")
st.write("Sube tu archivo `.docx`. El sistema validará tus rubros en cualquier formato (mayúsculas o minúsculas) y marcará los campos vacíos en el encabezado.")

# Definición de los rubros mandatorios (ahora incluye descripción del lugar)
RUBROS_BASE = [
    "planteamiento del problema", 
    "antecedente", 
    "estudio de campo", 
    "dirección", 
    "descripción del lugar",
    "observacion",    
    "consideracion",  
    "conclusion"      
]

archivo_subido = st.file_uploader("Elige tu archivo de Word", type=["docx"])

if archivo_subido is not None:
    st.info("🔄 Ejecutando auditoría de contenido... Por favor, espera.")
    
    doc = docx.Document(archivo_subido)
    
    texto_completo = ""
    errores_encabezado_cuenta = 0
    errores_congruencia_cuenta = 0
    
    # Variables de control de contradicciones
    tiene_ecatepec = False
    tiene_iztapalapa = False
    nombre_agente_inicio = ""

    # 1. ESCANEO Y MARCADO SEGURO DEL ENCABEZADO (HEADER)
    for seccion in doc.sections:
        header = seccion.header
        if header:
            texto_unificado_header = " ".join([p.text.lower() for p in header.paragraphs if p.text.strip()])
            tiene_datos_carpeta = any(caracter.isdigit() for caracter in texto_unificado_header)

            for parrafo in header.paragraphs:
                texto_linea = parrafo.text.strip()
                if not texto_linea:
                    continue
                texto_linea_lower = texto_linea.lower()

                # Ignorar títulos institucionales fijos
                if any(t in texto_linea_lower for t in ["agencia", "centro federal", "unidad de", "especialidad"]):
                    continue

                # Validar Folio Vacío
                if "folio" in texto_linea_lower:
                    contenido_folio = texto_linea.split(":")[-1].strip() if ":" in texto_linea else re.sub(r'número de folio|numero de folio', '', texto_linea, flags=re.IGNORECASE).strip()
                    if len(contenido_folio) == 0 and "[⚠️" not in texto_linea:
                        if parrafo.runs:
                            parrafo.runs[-1].text += " [⚠️ ERROR: FALTA LLENAR EL NÚMERO DE FOLIO]"
                        else:
                            parrafo.add_run(" [⚠️ ERROR: FALTA LLENAR EL NÚMERO DE FOLIO]")
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

                # Validar Carpeta Vacía
                elif "carpeta" in texto_linea_lower or "investigación" in texto_linea_lower or "investigacion" in texto_linea_lower:
                    if not tiene_datos_carpeta and "[⚠️" not in texto_linea:
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

    # 2. ESCANEO GLOBAL PREVIO EN EL CUERPO
    for parrafo in doc.paragraphs:
        txt_lower = parrafo.text.lower()
        texto_completo += " " + txt_lower
        if "ecatepec" in txt_lower:
            tiene_ecatepec = True
        if "iztapalapa" in txt_lower:
            tiene_iztapalapa = True
        if "maría del rocio" in txt_lower or "ramírez tentle" in txt_lower:
            if not nombre_agente_inicio:
                nombre_agente_inicio = txt_lower

    # 3. AUDITORÍA DE CONTRADICCIONES EN EL CUERPO
    for parrafo in doc.paragraphs:
        texto_original = parrafo.text
        texto_lower = texto_original.lower()
        if not texto_original.strip():
            continue

        # Marcar Contradicción Geográfica (Ecatepec vs Iztapalapa)
        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in texto_lower and "[⚠️" not in texto_original:
            if parrafo.runs:
                parrafo.runs[-1].text += " [⚠️ CONTRADICCIÓN: REVISAR MUNICIPIO EN PLANTILLA]"
            else:
                parrafo.add_run(" [⚠️ CONTRADICCIÓN: REVISAR MUNICIPIO EN PLANTILLA]")
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            errores_congruencia_cuenta += 1

    st.success("✅ ¡Auditoría de consistencia completada con éxito!")
    st.divider()

    # ------------------ REPORTES EN PANTALLA ------------------
    st.subheader("📊 Alertas de Contenido Detectadas")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Campos Vacíos Encabezado", errores_encabezado_cuenta)
    with col2:
        st.metric("Contradicciones de Plantilla", errores_congruencia_cuenta)
        
    if tiene_ecatepec and tiene_iztapalapa:
        st.error("⚠️ **Alerta Geográfica:** Se detectó el uso de 'Ecatepec' e 'Iztapalapa' simultáneamente en el texto.")

    # Verificar presencia de Rubros Obligatorios (Búsqueda flexible para minúsculas y acentos)
    st.subheader("📋 Control de Rubros Estructurados")
    rubros_faltantes = []
    
    # Normalizamos el texto completo eliminando acentos solo para la comparación de rubros
    texto_completo_limpio = texto_completo.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    
    for rubro in RUBROS_BASE:
        # Quitamos acentos al rubro base para buscarlo de forma limpia
        rubro_limpio = rubro.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        
        # Expresión regular tolerante a variantes en singular o plural
        patron = rf"{rubro_limpio}(es|s)?\b"
        if not re.search(patron, texto_completo_limpio):
            rubros_faltantes.append(rubro.upper())

    if rubros_faltantes:
        st.error(f"❌ Faltan los siguientes rubros obligatorios en el documento: {', '.join(rubros_faltantes)}")
    else:
        st.success("🎉 Todos los rubros mandatorios están presentes en el cuerpo del dictamen.")

    st.divider()

    # ------------------ BOTÓN DE DESCARGA ------------------
    st.subheader("📥 Descarga tu archivo corregido")
    st.write("Tu documento mantendrá intacta tu redacción original, mostrando únicamente marcas amarillas en los datos faltantes.")
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Documento Revisado",
        data=bio,
        file_name="DICTAMEN_REVISADO_IMPECABLE.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
