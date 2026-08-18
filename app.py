import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
from spellchecker import SpellChecker
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial Automatizado", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial Automatizado de Dictámenes")
st.write("Sube tu archivo de Word. El sistema generará marcas visuales directas dentro del documento para asegurar compatibilidad total con Microsoft Word.")

# Definición de los 7 rubros mandatorios
RUBROS_BASE = [
    "planteamiento del problema", "antecedente", "estudio de campo", 
    "dirección", "observacion", "consideracion", "conclusion"
]

# Diccionario de tecnicismos periciales/legales en México para evitar falsos positivos
PALABRAS_SEGURAS = {
    "siendo", "las", "direccion", "dirección", "perita", "perito", "adscrito", "adscrita",
    "exordio", "dictamen", "antecedentes", "planteamiento", "método", "técnica", "estudio",
    "gabinete", "observación", "observaciones", "consideración", "consideraciones", "conclusión",
    "conclusiones", "atentamente", "folio", "carpeta", "investigación", "expediente", "nuc",
    "cbtis", "insurgentes", "ecatepec", "iztapalapa", "comonfort", "maza", "parada", "tentle",
    "fgr", "aic", "pfm", "uinp", "diedcs", "sa"
}

archivo_subido = st.file_uploader("Elige tu archivo de Word", type=["docx"])

if archivo_subido is not None:
    st.info("🔄 Ejecutando auditoría y generando marcas visuales seguras... Por favor, espera.")
    
    doc = docx.Document(archivo_subido)
    spell = SpellChecker(language='es')
    
    texto_completo = ""
    errores_formato_cuenta = 0
    errores_ortografia_cuenta = 0
    errores_encabezado_cuenta = 0
    errores_congruencia_cuenta = 0
    
    # Variables de control de contradicciones
    tiene_ecatepec = False
    tiene_iztapalapa = False
    nombre_agente_inicio = ""

    # 1. ESCANEO Y MARCADO DEL ENCABEZADO (HEADER)
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

                if any(t in texto_linea_lower for t in ["agencia", "centro federal", "unidad de", "especialidad"]):
                    continue

                if "folio" in texto_linea_lower:
                    contenido_folio = texto_linea.split(":")[-1].strip() if ":" in texto_linea else re.sub(r'número de folio|numero de folio', '', texto_linea, flags=re.IGNORECASE).strip()
                    if len(contenido_folio) == 0 and "[⚠️" not in texto_linea:
                        if parrafo.runs:
                            parrafo.runs[-1].text += " [⚠️ ERROR DE CONTENIDO: FALTA LLENAR EL NÚMERO DE FOLIO]"
                        else:
                            parrafo.add_run(" [⚠️ ERROR DE CONTENIDO: FALTA LLENAR EL NÚMERO DE FOLIO]")
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

                elif "carpeta" in texto_linea_lower or "investigación" in texto_linea_lower:
                    if not tiene_datos_carpeta and "[⚠️" not in texto_linea:
                        if parrafo.runs:
                            parrafo.runs[-1].text += " [⚠️ ERROR DE CONTENIDO: FALTA LLENAR LA CARPETA DE INVESTIGACIÓN]"
                        else:
                            parrafo.add_run(" [⚠️ ERROR DE CONTENIDO: FALTA LLENAR LA CARPETA DE INVESTIGACIÓN]")
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

    # 3. AUDITORÍA DEL CUERPO (MARCAS VISUALES SEGURAS)
    for parrafo in doc.paragraphs:
        texto_original = parrafo.text
        texto_lower = texto_original.lower()
        if not texto_original.strip():
            continue

        # --- AUDITORÍA DE CONTENIDO: CONTRADICCIÓN GEOGRÁFICA ---
        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in texto_lower and "[⚠️" not in texto_original:
            if parrafo.runs:
                parrafo.runs[-1].text += " [⚠️ CONTRADICCIÓN: Mencionas Iztapalapa aquí, pero en el Estudio de Campo declaraste Ecatepec.]"
            else:
                parrafo.add_run(" [⚠️ CONTRADICCIÓN: Mencionas Iztapalapa aquí, pero en el Estudio de Campo declaraste Ecatepec.]")
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            errores_congruencia_cuenta += 1

        # --- AUDITORÍA DE CONTENIDO: ALTERACIÓN DE NOMBRES ---
        if "ramírez" in texto_lower and "maría" in texto_lower and "[⚠️" not in texto_original:
            if "ramírez tentle maría" in texto_lower and "maría del rocio" in nombre_agente_inicio:
                if parrafo.runs:
                    parrafo.runs[-1].text += " [⚠️ ALERTA: El orden de los apellidos de la autoridad cambió respecto al exordio.]"
                else:
                    parrafo.add_run(" [⚠️ ALERTA: El orden de los apellidos de la autoridad cambió respecto al exordio.]")
                errores_congruencia_cuenta += 1

        # --- REVISIÓN ORTOGRÁFICA VISUALMENTE COMPATIBLE ---
        palabras = re.findall(r'\b\w+\b', texto_original)
        for palabra in palabras:
            palabra_lower = palabra.lower()
            if len(palabra_lower) > 2 and not palabra_lower.isdigit() and palabra_lower not in PALABRAS_SEGURAS:
                if not spell.known([palabra_lower]):
                    sugerencia = spell.correction(palabra)
                    if sugerencia and sugerencia.lower() != palabra_lower:
                        # Buscamos el fragmento (run) que contiene la palabra y la resaltamos
                        for run in parrafo.runs:
                            if palabra in run.text:
                                run.text = run.text.replace(palabra, f"{sugerencia} (antes: {palabra})")
                                run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE # Resaltado color turquesa/azul
                                errores_ortografia_cuenta += 1

    st.success("✅ ¡Auditoría completada de forma segura para Microsoft Word!")
    st.divider()

    # Verificar presencia de Rubros Obligatorios
    st.subheader("📋 Estado de Rubros Obligatorios")
    rubros_faltantes = []
    texto_completo_limpio = texto_completo.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    for rubro in RUBROS_BASE:
        patron = rf"{rubro}(es|s)?\b"
        if not re.search(patron, texto_completo_limpio):
            rubros_faltantes.append(rubro.upper())

    if rubros_faltantes:
        st.error(f"❌ Faltan los siguientes rubros en el cuerpo: {', '.join(rubros_faltantes)}")
    else:
        st.success("🎉 Todos los rubros obligatorios están presentes en el documento.")

    st.divider()

    # ------------------ BOTÓN DE DESCARGA ------------------
    st.subheader("📥 Descarga tu archivo auditado")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Documento Corregido",
        data=bio,
        file_name="DICTAMEN_REVISADO_SEGURO.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
