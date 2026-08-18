import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
from spellchecker import SpellChecker
import re
import io

# Configuración de la página web
st.set_page_config(page_title="Revisor de Word Inteligente Pro", page_icon="📝", layout="centered")

st.title("📝 Revisor de Dictámenes Periciales")
st.write("Sube tu archivo `.docx`. El sistema validará tus rubros obligatorios, la ortografía real y forzará marcas amarillas si el Folio o la Carpeta del encabezado están vacíos.")

# Lista de rubros base actualizada (flexibles a singular o plural - sin 'siendo las')
RUBROS_BASE = [
    "planteamiento del problema", 
    "antecedente", 
    "estudio de campo", 
    "dirección",
    "observacion",    
    "consideracion",  
    "conclusion"      
]

# Diccionario de palabras seguras comunes en el ámbito legal/pericial para evitar falsos positivos
PALABRAS_SEGURAS = {
    "siendo", "las", "direccion", "dirección", "perita", "perito", "adscrito", "adscrita",
    "exordio", "dictamen", "antecedentes", "planteamiento", "método", "técnica", "estudio",
    "gabinete", "observación", "observaciones", "consideración", "consideraciones", "conclusión",
    "conclusiones", "atentamente", "folio", "carpeta", "investigación", "expediente", "nuc"
}

archivo_subido = st.file_uploader("Elige tu archivo de Word", type=["docx"])

if archivo_subido is not None:
    st.info("🔄 Analizando y verificando tu documento... Por favor, espera.")
    
    # Leer el documento original
    doc = docx.Document(archivo_subido)
    spell = SpellChecker(language='es')
    
    texto_completo = ""
    errores_formato_cuenta = 0
    errores_ortografia_cuenta = 0
    errores_encabezado_cuenta = 0

    # 1. REVISIÓN Y MARCADO DEL ENCABEZADO (HEADER)
    for seccion in doc.sections:
        header = seccion.header
        if header:
            for parrafo in header.paragraphs:
                texto_linea = parrafo.text.strip()
                if not texto_linea:
                    continue

                texto_linea_lower = texto_linea.lower()

                # Candado de seguridad: Ignorar los títulos oficiales fijos de la institución
                if any(t in texto_linea_lower for t in ["agencia", "centro federal", "unidad de", "especialidad"]):
                    continue

                # --- VALIDACIÓN 1: NÚMERO DE FOLIO ---
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

                # --- VALIDACIÓN 2: CARPETA DE INVESTIGACIÓN ---
                elif "carpeta" in texto_linea_lower or "investigación" in texto_linea_lower or "investigacion" in texto_linea_lower:
                    tiene_numeros = any(caracter.isdigit() for caracter in texto_linea)
                    
                    if not tiene_numeros:
                        # Forzar el marcador amarillo únicamente sobre la línea de la carpeta vacía
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

    # 2. ANALIZAR EL CUERMO DEL DOCUMENTO
    for parrafo in doc.paragraphs:
        texto_parrafo_limpio = parrafo.text.lower()
        texto_parrafo_limpio = texto_parrafo_limpio.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        texto_completo += " " + texto_parrafo_limpio

    # Tercera pasada: Control de diseño y ortografía real en el cuerpo
    for i, parrafo in enumerate(doc.paragraphs, start=1):
        texto_original = parrafo.text
        if not texto_original.strip():
            continue

        # Control de diseño general (Raleway 9 a 11)
        for run in parrafo.runs:
            if run.text.strip():
                fuente = run.font.name
                tamaño = run.font.size.pt if run.font.size else None
                
                fuente_incorrecta = fuente and fuente != "Raleway"
                tamaño_incorrecto = tamaño and (tamaño < 9.0 or tamaño > 11.0)
                
                if fuente_incorrecta or tamaño_incorrecto:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    errores_formato_cuenta += 1

        # Conteo de ORTOGRAFÍA limpia (ignora mayúsculas/minúsculas y palabras seguras)
        palabras = re.findall(r'\b\w+\b', texto_original)
        for palabra in palabras:
            palabra_lower = palabra.lower()
            if len(palabra_lower) > 2 and not palabra_lower.isdigit():
                if palabra_lower not in PALABRAS_SEGURAS:
                    if not spell.known([palabra_lower]):
                        errores_ortografia_cuenta += 1

    st.success("✅ ¡Revisión completada!")
    st.divider()
    
    # ------------------ REPORTES EN PANTALLA ------------------
    st.subheader("📊 Resumen de Alertas en el Documento")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Campos Vacíos en Encabezado", errores_encabezado_cuenta)
    with col2:
        st.metric("Alertas de Formato (Raleway 9-11)", errores_formato_cuenta)
    with col3:
        st.metric("Palabras con Ortografía Dudosa", errores_ortografia_cuenta)
        
    # Verificar Rubros
    st.subheader("📋 Estado de Rubros Obligatorios")
    rubros_faltantes = []
    for rubro in RUBROS_BASE:
        if rubro in ["observacion", "consideracion", "conclusion"]:
            patron = rf"{rubro}(es|s)?\b"
            if not re.search(patron, texto_completo):
                rubros_faltantes.append(f"{rubro} / {rubro}es")
        else:
            patron = rf"{rubro}s?\b"
            if not re.search(patron, texto_completo):
                rubros_faltantes.append(rubro)
    
    if rubros_faltantes:
        st.error(f"❌ Faltan o están mal escritos {len(rubros_faltantes)} rubros obligatorios:")
        for rf in rubros_faltantes:
            st.markdown(f"* **{rf.upper()}**")
    else:
        st.success("🎉 ¡Excelente! El documento contiene todos tus rubros obligatorios.")
        
    st.divider()
    
    # ------------------ BOTÓN DE DESCARGA ------------------
    st.subheader("📥 Descarga tu archivo marcado")
    st.write("Abre el archivo descargado en Word para visualizar el folio y la carpeta marcados en amarillo.")
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Documento Revisado",
        data=bio,
        file_name="DICTAMEN_REVISADO_ESTRICTO.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
