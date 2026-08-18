import streamlit as st
import docx
from docx.enum.text import WD_COLOR_INDEX
from spellchecker import SpellChecker
import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial de Dictámenes", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial de Dictámenes")
st.write("Sube tu archivo `.docx`. El sistema validará rubros y contradicciones directamente en el Word, y desplegará el reporte ortográfico en la pantalla web.")

# Definición de los rubros mandatorios
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

# Diccionario de tecnicismos y palabras autorizadas para evitar falsos positivos
PALABRAS_SEGURAS = {
    "siendo", "las", "direccion", "dirección", "perita", "perito", "adscrito", "adscrita",
    "exordio", "dictamen", "antecedentes", "planteamiento", "método", "técnica", "estudio",
    "gabinete", "observación", "observaciones", "consideración", "consideraciones", "conclusión",
    "conclusiones", "atentamente", "folio", "carpeta", "investigación", "expediente", "nuc",
    "cbtis", "insurgentes", "ecatepec", "iztapalapa", "comonfort", "maza", "parada", "tentle",
    "fgr", "aic", "pfm", "uinp", "diedcs", "sa", "criminalística", "criminalistica", "designe"
}

archivo_subido = st.file_uploader("Elige tu archivo de Word", type=["docx"])

if archivo_subido is not None:
    st.info("🔄 Ejecutando auditoría de contenido y ortografía... Por favor, espera.")
    
    doc = docx.Document(archivo_subido)
    spell = SpellChecker(language='es')
    
    texto_completo = ""
    errores_encabezado_cuenta = 0
    errores_congruencia_cuenta = 0
    palabras_sospechosas = []

    # 1. REVISIÓN Y MARCADO SEGURO DEL ENCABEZADO (HEADER)
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
                    if not tiene_datos_carpeta and "[⚠️" not in texto_linea:
                        for run in parrafo.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        errores_encabezado_cuenta += 1

    # 2. ESCANEO GLOBAL PREVIO EN EL CUERPO (ANÁLISIS ORTOGRÁFICO EN PANTALLA)
    tiene_ecatepec = False
    tiene_iztapalapa = False

    for i, parrafo in enumerate(doc.paragraphs, start=1):
        txt = parrafo.text.strip()
        if not txt:
            continue
        txt_lower = txt.lower()
        texto_completo += " " + txt_lower
        
        if "ecatepec" in txt_lower:
            tiene_ecatepec = True
        if "iztapalapa" in txt_lower:
            tiene_iztapalapa = True

        # --- REVISIÓN ORTOGRÁFICA TRATADA COMO TEXTO SEGURO ---
        palabras = re.findall(r'\b\w+\b', txt)
        for palabra in palabras:
            palabra_lower = palabra.lower()
            if len(palabra_lower) > 2 and not palabra_lower.isdigit() and palabra_lower not in PALABRAS_SEGURAS:
                # Comprobación de acentos estrictos (como Roció vs Rocío)
                if not spell.known([palabra]) or palabra_lower == "roció":
                    sugerencia = "Rocío" if palabra_lower == "roció" else spell.correction(palabra_lower)
                    if sugerencia and sugerencia.lower() != palabra_lower:
                        palabras_sospechosas.append({"parrafo": i, "incorrecta": palabra, "sugerencia": sugerencia})

    # 3. AUDITORÍA DE CONTRADICCIONES EN EL CUERPO (MARCADO AMARILLO SEGURO)
    for parrafo in doc.paragraphs:
        texto_original = parrafo.text
        texto_lower = texto_original.lower()
        if not texto_original.strip():
            continue

        if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in texto_lower and "[⚠️" not in texto_original:
            if parrafo.runs:
                parrafo.runs[-1].text += " [⚠️ CONTRADICCIÓN: REVISAR MUNICIPIO EN PLANTILLA]"
            else:
                parrafo.add_run(" [⚠️ CONTRADICCIÓN: REVISAR MUNICIPIO EN PLANTILLA]")
            for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            errores_congruencia_cuenta += 1

    st.success("✅ ¡Auditoría e inspección ortográfica completadas!")
    st.divider()

    # ------------------ REPORTE EN PANTALLA WEB: ORTOGRAFÍA ------------------
    st.subheader("📝 1. Reporte de Ortografía y Acentuación")
    if palabras_sospechosas:
        st.warning(f"Se detectaron {len(palabras_sospechosas)} palabras sospechosas en el texto libre:")
        vistas = set()
        for item in palabras_sospechosas:
            clave = f"{item['incorrecta']}->{item['sugerencia']}"
            if clave not in vistas:
                st.markdown(f"* En Párrafo **{item['parrafo']}** dice: **\"{item['incorrecta']}\"** 👉 ¿Quisiste decir: *{item['sugerencia']}*?")
                vistas.add(clave)
    else:
        st.success("🎉 ¡Excelente! No se detectaron faltas de ortografía evidentes en el texto.")

    st.divider()

    # REPORTES EN PANTALLA DE CONTENIDO
    st.subheader("📊 2. Alertas de Contenido Detectadas")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Campos Vacíos Encabezado", errores_encabezado_cuenta)
    with col2:
        st.metric("Contradicciones de Plantilla", errores_congruencia_cuenta)

    # Verificar presencia de Rubros Obligatorios
    st.subheader("📋 3. Control de Rubros Estructurados")
    rubros_faltantes = []
    texto_completo_limpio = texto_completo.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    
    for rubro in RUBROS_BASE:
        rubro_limpio = rubro.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        patron = rf"{rubro_limpio}(es|s)?\b"
        if not re.search(patron, texto_completo_limpio):
            rubros_faltantes.append(rubro.upper())

    if rubros_faltantes:
        st.error(f"❌ Faltan los siguientes rubros obligatorios en el documento: {', '.join(rubros_faltantes)}")
    else:
        st.success("🎉 Todos los rubros mandatorios están presentes en el cuerpo del dictamen.")

    st.divider()

    # ------------------ BOTÓN DE DESCARGA INTEGRAL Y SEGURO ------------------
    st.subheader("📥 Descarga tu archivo marcado")
    st.write("El archivo mantendrá intacta tu redacción original, iluminando únicamente en amarillo los campos vacíos y contradicciones.")
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Documento Revisado",
        data=bio,
        file_name="DICTAMEN_REVISADO_IMPECABLE.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
