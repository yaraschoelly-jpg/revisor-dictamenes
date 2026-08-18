
import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
from spellchecker import SpellChecker
import re
import io

# Configuración de la página web
st.set_page_config(page_title="Revisor de Word Inteligente Pro", page_icon="📝", layout="centered")

st.title("📝 Revisor de Dictámenes Periciales Pro")
st.write("Prototipo avanzado: Validación de formato, rubros, ortografía e incongruencias globales de contenido con explicaciones.")

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
    st.info("🔄 Ejecutando auditoría pericial global y cruzando datos... Por favor, espera.")
    
    doc = docx.Document(archivo_subido)
    spell = SpellChecker(language='es')
    
    texto_completo = ""
    errores_formato_cuenta = 0
    errores_ortografia_cuenta = 0
    errores_encabezado_cuenta = 0
    errores_congruencia_cuenta = 0
    
    # Variables de control global para contradicciones
    tiene_ecatepec = False
    tiene_iztapalapa = False
    nombre_agente_inicio = ""
    
    # Listas para almacenar mensajes explicativos detallados
    motivos_subrayado = []

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
                        if parrafo.runs:
                            parrafo.runs[-1].text += " [⚠️ ERROR: FALTA LLENAR EL NÚMERO DE FOLIO]"
                        else:
                            parrafo.add_run(" [⚠️ ERROR: FALTA LLENAR EL NÚMERO DE FOLIO]")
                        
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

                elif "carpeta" in texto_linea_lower or "investigación" in texto_linea_lower or "investigacion" in texto_linea_lower:
                    tiene_numeros = any(caracter.isdigit() for caracter in texto_linea)
                    if not tiene_numeros:
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

    # 2. PRIMERA PASADA: DETECCIÓN GLOBAL DE CONTRADICCIONES DE TEXTO
    for parrafo in doc.paragraphs:
        txt_lower = parrafo.text.lower()
        texto_completo += " " + txt_lower
        
        if "ecatepec" in txt_lower:
            tiene_ecatepec = True
        if "iztapalapa" in txt_lower:
            tiene_iztapalapa = True
            
        if "maría del rocio maritza" in txt_lower or "ramírez tentle" in txt_lower:
            if not nombre_agente_inicio:
                nombre_agente_inicio = txt_lower

    # 3. SEGUNDA PASADA: APLICAR RESALTADOS DIRECTOS Y REGISTRAR MOTIVOS EXPLICATIVOS
    for i, parrafo in enumerate(doc.paragraphs, start=1):
        texto_original = parrafo.text
        texto_lower = texto_original.lower()
        if not texto_original.strip():
            continue

        # --- VALIDACIÓN GLOBAL 1: INCONGRUENCIA GEOGRÁFICA ---
        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in texto_lower:
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            errores_congruencia_cuenta += 1
            # Añadir explicación con fragmento de texto real
            motivos_subrayado.append(f"❌ **Subrayado por Contradicción Geográfica:** El párrafo que dice *\"{texto_original[:90]}...\"* menciona **Iztapalapa**, lo cual contradice la ubicación de **Ecatepec** declarada en el cuerpo del Dictamen.")

        # --- VALIDACIÓN GLOBAL 2: MUTACIÓN DE NOMBRE DE AUTORIDAD ---
        if "ramírez" in texto_lower or "tentle" in texto_lower:
            if "ramírez tentle maría" in texto_lower and "maría del rocio" in nombre_agente_inicio:
                for run in parrafo.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                errores_congruencia_cuenta += 1
                motivos_subrayado.append(f"❌ **Subrayado por Alteración de Identidad:** En el párrafo *\"{texto_original[:90]}...\"* cambiaste el orden de los apellidos de la autoridad respecto a cómo la presentaste originalmente en el exordio.")

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
    
    # ------------------ NUEVO APARTADO VISUAL: EXPLICACIÓN DE SUBRAYADOS ------------------
    st.subheader("🕵️‍♂️ 1. Motivos de los Subrayados Amarillos")
    
    if motivos_subrayado:
        st.warning(f"Se realizaron {len(motivos_subrayado)} marcados por incongruencia de contenido:")
        for motivo in set(motivos_subrayado): # Evitar duplicados en pantalla
            st.markdown(motivo)
    else:
        st.success("🎉 ¡Excelente! No fue necesario subrayar párrafos por contradicciones de redacción.")

    st.divider()

    # REPORTES NUMÉRICOS EN PANTALLA
    st.subheader("📊 Resumen de Alertas Generales")
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
