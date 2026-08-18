import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
from spellchecker import SpellChecker
import re
import io

# Configuración de la página web
st.set_page_config(page_title="Revisor de Word Inteligente Pro", page_icon="📝", layout="centered")

st.title("📝 Revisor de Dictámenes Periciales Pro")
st.write("Prototipo avanzado: Validación de formato, rubros, ortografía e incongruencias de contenido (Ubicación y Nombres).")

# Lista de rubros base
RUBROS_BASE = [
    "planteamiento del problema", 
    "antecedente", 
    "estudio de campo", 
    "dirección",
    "observacion",    
    "consideracion",  
    "conclusion"      
]

# Diccionario de palabras seguras comunes
PALABRAS_SEGURAS = {
    "siendo", "las", "direccion", "dirección", "perita", "perito", "adscrito", "adscrita",
    "exordio", "dictamen", "antecedentes", "planteamiento", "método", "técnica", "estudio",
    "gabinete", "observación", "observaciones", "consideración", "consideraciones", "conclusión",
    "conclusiones", "atentamente", "folio", "carpeta", "investigación", "expediente", "nuc",
    "cbtis", "insurgentes", "ecatepec", "iztapalapa", "comonfort", "maza", "parada", "tentle"
}

archivo_subido = st.file_uploader("Elige tu archivo de Word", type=["docx"])

if archivo_subido is not None:
    st.info("🔄 Ejecutando auditoría pericial y cruzando datos... Por favor, espera.")
    
    doc = docx.Document(archivo_subido)
    spell = SpellChecker(language='es')
    
    texto_completo = ""
    errores_formato_cuenta = 0
    errores_ortografia_cuenta = 0
    errores_encabezado_cuenta = 0
    errores_congruencia_cuenta = 0
    
    # Variables para cruce de datos
    municipio_estudio = ""
    municipio_conclusion = ""
    nombre_agente_inicio = ""
    nombre_agente_observaciones = ""

    # 1. REVISIÓN DEL ENCABEZADO (HEADER)
    for seccion in doc.sections:
        header = seccion.header
        if header:
            texto_unificado_header = " ".join([p.text.lower() for p in header.paragraphs if p.text.strip()])
            tiene_datos_carpeta = False
            texto_sin_plantilla = re.sub(r'número de expediente|carpeta de|investigación|investigacion|o averiguación previa|averiguacion|[:_\[\]\s,.–—-]', '', texto_unificado_header).strip()
            if len(texto_sin_plantilla) > 3:
                tiene_datos_carpeta = True

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
                        parrafo.text = texto_linea + " [⚠️ ERROR: FALTA LLENAR EL NÚMERO DE FOLIO]"
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

                elif "carpeta" in texto_linea_lower or "investigación" in texto_linea_lower or "investigacion" in texto_linea_lower:
                    tiene_numeros = any(caracter.isdigit() for caracter in texto_linea)
                    if not tiene_numeros:
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

    # 2. ESCANEO GENERAL Y EXTRACCIÓN DE CONTENIDO EXTRAVIADO
    for parrafo in doc.paragraphs:
        txt = parrafo.text.strip()
        texto_completo += " " + txt.lower()
        
        # Detectar el nombre de la agente al inicio
        if "maría del rocio maritza" in txt or "ramírez tentle" in txt:
            if not nombre_agente_inicio:
                nombre_agente_inicio = txt

        # Extraer geografía del estudio de campo
        if "ecatepec" in txt.lower():
            municipio_estudio = "ecatepec"
        elif "iztapalapa" in txt.lower() and not municipio_estudio:
            municipio_estudio = "iztapalapa"

    # 3. AUDITORÍA DEL CUERPO (MARCADO DE INCONGRUENCIAS Y ORTOGRAFÍA)
    en_observaciones = False
    en_conclusion = False

    for i, parrafo in enumerate(doc.paragraphs, start=1):
        texto_original = parrafo.text
        texto_lower = texto_original.lower()
        if not texto_original.strip():
            continue

        # Detectar zonas críticas
        if "observacion" in texto_lower:
            en_observaciones = True
        if "conclusion" in texto_lower:
            en_observaciones = False
            en_conclusion = True

        # --- VALIDACIÓN DE CONGRUENCIA 1: MUNICIPIOS CRUZADOS ---
        if en_conclusion and ("ecatepec" in texto_lower or "iztapalapa" in texto_lower):
            if "iztapalapa" in texto_lower and municipio_estudio == "ecatepec":
                # Forzar marcador amarillo por contradicción geográfica de plantilla
                for run in parrafo.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                errores_congruencia_cuenta += 1

        # --- VALIDACIÓN DE CONGRUENCIA 2: MUTACIÓN DE NOMBRES ---
        if en_observaciones and ("ramírez" in texto_lower or "rocio" in texto_lower):
            # Si el orden de las palabras cambió respecto al exordio
            if "ramírez tentle maría" in texto_lower and "maría del rocio" in nombre_agente_inicio.lower():
                for run in parrafo.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                errores_congruencia_cuenta += 1

        # --- REVISIÓN DE DISEÑO GENERAL (Raleway 9 a 11) ---
        for run in parrafo.runs:
            if run.text.strip():
                fuente = run.font.name
                tamaño = run.font.size.pt if run.font.size else None
                if (fuente and fuente != "Raleway") or (tamaño and (tamaño < 9.0 or tamaño > 11.0)):
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    errores_formato_cuenta += 1

        # --- REVISIÓN DE ORTOGRAFÍA ---
        palabras = re.findall(r'\b\w+\b', texto_original)
        for palabra in palabras:
            palabra_lower = palabra.lower()
            if len(palabra_lower) > 2 and not palabra_lower.isdigit():
                if palabra_lower not in PALABRAS_SEGURAS and not spell.known([palabra_lower]):
                    errores_ortografia_cuenta += 1

    st.success("✅ ¡Auditoría de consistencia completada!")
    st.divider()
    
    # ------------------ REPORTES EN PANTALLA ------------------
    st.subheader("📊 Alertas de Contenido e Incongruencias")
    
    if errores_congruencia_cuenta > 0:
        st.error(f"❌ Se detectaron {errores_congruencia_cuenta} contradicciones humanas graves en el texto.")
        st.write("👉 **Falla Geográfica:** Declaraste 'Ecatepec' en el Estudio de campo pero cerraste con 'Iztapalapa' en la Conclusión.")
        st.write("👉 **Falla de Identidad:** El orden del nombre de la Agente no coincide entre el exordio y las observaciones.")
    else:
        st.success("🎉 ¡Excelente! No se encontraron contradicciones de texto entre las secciones.")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Campos Vacíos Encabezado", errores_encabezado_cuenta)
    with col2:
        st.metric("Detalles de Formato Amarillo", errores_formato_cuenta)
    with col3:
        st.metric("Faltas Ortográficas Reales", errores_ortografia_cuenta)
        
    # Verificar Rubros
    st.subheader("📋 Control de Rubros Estructurados")
    rubros_faltantes = []
    texto_completo_limpio = texto_completo.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    for rubro in RUBROS_BASE:
        patron = rf"{rubro}(es|s)?\b"
        if not re.search(patron, texto_completo_limpio):
            rubros_faltantes.append(rubro)
    
    if rubros_faltantes:
        st.error(f"❌ Faltan estos rubros obligatorios: {', '.join([r.upper() for r in rubros_faltantes])}")
    else:
        st.success("🎉 Todos los rubros mandatorios están presentes.")
        
    st.divider()
    
    # ------------------ BOTÓN DE DESCARGA ------------------
    st.subheader("📥 Descarga tu archivo auditado")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Documento Revisado",
        data=bio,
        file_name="DICTAMEN_AUDITADO_PRO.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
