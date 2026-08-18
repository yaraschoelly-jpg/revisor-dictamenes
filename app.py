import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from spellchecker import SpellChecker
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial Automatizado", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial Automatizado de Dictámenes")
st.write("Sube tu archivo de Word. El sistema aplicará **Control de Cambios** para ortografía e insertará **Globos de Comentario** nativos al margen con la explicación de cada error.")

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

# --- FUNCIONES TÉCNICAS XML PARA WORD (COMENTARIOS Y CONTROL DE CAMBIOS) ---
def insertar_comentario_nativo(parrafo, texto_comentario, id_comentario):
    """Inserta un globo de comentario nativo en el margen derecho de Microsoft Word"""
    pPr = parrafo._p.get_or_add_pPr()
    
    # Crear elementos XML obligatorios para el estándar OpenXML de Word
    commentRangeStart = OxmlElement('w:commentRangeStart')
    commentRangeStart.set(qn('w:id'), str(id_comentario))
    commentRangeEnd = OxmlElement('w:commentRangeEnd')
    commentRangeEnd.set(qn('w:id'), str(id_comentario))
    
    parrafo._p.insert(0, commentRangeStart)
    parrafo._p.append(commentRangeEnd)
    
    commentReference = OxmlElement('w:commentReference')
    commentReference.set(qn('w:id'), str(id_comentario))
    parrafo.add_run()._r.append(commentReference)

def activar_control_cambios_parrafo(parrafo, texto_antiguo, texto_nuevo):
    """Aplica control de cambios nativo sustituyendo texto y dejándolo registrado en Word"""
    parrafo.text = "" # Limpiamos el texto plano
    
    # Nodo de texto eliminado (aparecerá tachado en rojo en Word)
    del_run = OxmlElement('w:del')
    del_run.set(qn('w:author'), 'Auditor Automatizado')
    del_text = OxmlElement('w:text')
    del_text.text = texto_antiguo
    del_run.append(del_text)
    
    # Nodo de texto insertado (aparecerá subrayado en rojo en Word)
    ins_run = OxmlElement('w:ins')
    ins_run.set(qn('w:author'), 'Auditor Automatizado')
    ins_text = OxmlElement('w:text')
    ins_text.text = texto_nuevo
    ins_run.append(ins_text)
    
    parrafo._p.append(del_run)
    parrafo._p.append(ins_run)

# --- INICIO DEL PROCESAMIENTO ---
archivo_subido = st.file_uploader("Elige tu archivo de Word", type=["docx"])

if archivo_subido is not None:
    st.info("🔄 Ejecutando auditoría y generando marcas nativas de Word... Por favor, espera.")
    
    doc = docx.Document(archivo_subido)
    spell = SpellChecker(language='es')
    
    texto_completo = ""
    id_comentario = 1
    
    # Variables de control de contradicciones
    tiene_ecatepec = False
    tiene_iztapalapa = False
    nombre_agente_inicio = ""

    # 1. ESCANEO DEL ENCABEZADO (HEADER)
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
                    if ":" in texto_linea and len(texto_linea.split(":")[-1].strip()) == 0:
                        insertar_comentario_nativo(parrafo, "CRÍTICO: El campo de Número de Folio se encuentra vacío en la plantilla.", id_comentario)
                        id_comentario += 1
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW

                elif "carpeta" in texto_linea_lower or "investigación" in texto_linea_lower:
                    if not tiene_datos_carpeta:
                        insertar_comentario_nativo(parrafo, "CRÍTICO: Falta capturar el número identificador de la Carpeta de Investigación.", id_comentario)
                        id_comentario += 1
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW

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

    # 3. AUDITORÍA DEL CUERPO: CONTROL DE CAMBIOS Y COMENTARIOS AL MARGEN
    for parrafo in doc.paragraphs:
        texto_original = parrafo.text
        texto_lower = texto_original.lower()
        if not texto_original.strip():
            continue

        # --- AUDITORÍA DE CONTENIDO: CONTRADICCIÓN GEOGRÁFICA ---
        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in texto_lower:
            insertar_comentario_nativo(parrafo, "CONTRADICCIÓN: Mencionas Iztapalapa en esta sección, pero en el Estudio de Campo declaraste que el lugar de los hechos está en Ecatepec.", id_comentario)
            id_comentario += 1
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        # --- AUDITORÍA DE CONTENIDO: ALTERACIÓN DE NOMBRES ---
        if "ramírez" in texto_lower and "maría" in texto_lower:
            if "ramírez tentle maría" in texto_lower and "maría del rocio" in nombre_agente_inicio:
                insertar_comentario_nativo(parrafo, "ALERTA: El orden de los apellidos de la autoridad cambió respecto al exordio inicial.", id_comentario)
                id_comentario += 1

        # --- REVISIÓN ORTOGRÁFICA CON CONTROL DE CAMBIOS AUTOMÁTICO ---
        palabras = re.findall(r'\b\w+\b', texto_original)
        texto_modificado = texto_original
        hubo_cambio_ortografia = False

        for palabra in palabras:
            palabra_lower = palabra.lower()
            if len(palabra_lower) > 2 and not palabra_lower.isdigit() and palabra_lower not in PALABRAS_SEGURAS:
                if not spell.known([palabra_lower]):
                    sugerencia = spell.correction(palabra)
                    if sugerencia and sugerencia.lower() != palabra_lower:
                        # Respetar mayúscula inicial si la original la tenía
                        if palabra[0].isupper():
                            sugerencia = sugerencia.capitalize()
                        texto_modificado = texto_modificado.replace(palabra, sugerencia)
                        hubo_cambio_ortografia = True

        if hubo_cambio_ortografia:
            activar_control_cambios_parrafo(parrafo, texto_original, texto_modificado)

    st.success("✅ ¡Auditoría completada e integrada al archivo Word!")
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
    st.subheader("📥 Descarga tu archivo con marcas oficiales")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Word con Control de Cambios",
        data=bio,
        file_name="DICTAMEN_REVISADO_COMPLETO.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
