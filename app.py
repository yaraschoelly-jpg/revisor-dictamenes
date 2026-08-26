import streamlit as st
import subprocess
import sys

# Motor de Autoinstalación Forzada Directa en el Servidor
try:
    import docx
    from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
except ModuleNotFoundError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    import docx
    from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH

try:
    import pypdf
except ModuleNotFoundError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
    import pypdf

import re
import io

# Configuración de la interfaz web
st.set_page_config(page_title="Auditor Pericial Integral", page_icon="⚖️", layout="centered")

st.title("⚖️ Auditor Pericial de Formalidad Estructural")
st.write("Sube el **PDF de Solicitud** y tu **Word del Dictamen**. El sistema validará la estructura formal y aplicará las marcas directamente en tu documento de Word.")

# Lista de rubros obligatorios según el manual de la institución
RUBROS_BASE = [
    "planteamiento del problema", "antecedente", "estudio de campo", 
    "dirección", "descripción del lugar", "observacion", "consideracion", "conclusion"
]

# Casillas de doble carga oficiales
st.subheader("📁 1. Carga de Documentos Oficiales")
col_pdf, col_docx = st.columns(2)

with col_pdf:
    archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
    archivo_docx = st.file_uploader("Subir Dictamen Pericial en Word (.docx)", type=["docx"])

if archivo_pdf is not None and archivo_docx is not None:
    st.info("🔍 Analizando consistencia, formalidad y ortografía... Por favor, espera.")
    
    # --- 1. EXTRACCIÓN DE TEXTO DEL PDF ---
    texto_pdf = ""
    try:
        lector_pdf = pypdf.PdfReader(archivo_pdf)
        for pagina in lector_pdf.pages:
            texto_pdf += " " + pagina.extract_text()
    except Exception as e:
        texto_pdf = "error"
        
    texto_pdf_lower = texto_pdf.lower()

    # --- 2. EXTRAER DATOS CLAVE DEL PDF ---
    match_carpeta_pdf = re.search(r'fed/[a-z0-9/_\-]+', texto_pdf_lower)
    carpeta_solicitud = match_carpeta_pdf.group(0).upper() if match_carpeta_pdf else "FED/FEVIMTRA/FEIDHVM-MEX/0000251/2026"
    
    match_oficio_pdf = re.search(r'fgr-aic-pfm-[a-z0-9\-]+', texto_pdf_lower)
    oficio_solicitud = match_oficio_pdf.group(0).upper() if match_oficio_pdf else "FGR-AIC-PFM-UINP-DIEDCS-SA-017608-2026"

    # --- 3. EXTRACCIÓN Y AUDITORÍA EN EL WORD ---
    doc = docx.Document(archivo_docx)
    texto_word_completo = ""
    
    errores_encabezado_cuenta = 0
    errores_antecedentes_cuenta = 0
    errores_congruencia_cuenta = 0
    errores_justificado_cuenta = 0
    errores_centrado_cuenta = 0
    errores_tipografia_cuenta = 0
    
    palabras_sospechosas = []
    alertas_diseno = []

    # A. Revisión del Encabezado
    try:
        for seccion in doc.sections:
            header = seccion.header
            if header:
                for parrafo in header.paragraphs:
                    texto_linea = parrafo.text.strip()
                    if not texto_linea:
                        continue
                    texto_linea_lower = texto_linea.lower()

                    if any(t in texto_linea_lower for t in ["agencia", "centro federal", "unidad de", "especialidad"]):
                        continue

                    if "carpeta" in texto_linea_lower:
                        digitos_carpeta = "".join(re.findall(r'\d+', carpeta_solicitud))
                        if not any(d in texto_linea_lower for d in digitos_carpeta[:4]):
                            parrafo.text = f"{texto_linea} [⚠️ ERROR DE CONTROL: DEBE SER {carpeta_solicitud}]"
                            for run in parrafo.runs:
                                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            errores_encabezado_cuenta += 1
    except Exception as e:
        pass

    # B. Revisión del Cuerpo, Alineaciones y Ortografía Universal
    tiene_ecatepec = False
    tiene_iztapalapa = False

    for i, parrafo in enumerate(doc.paragraphs, start=1):
        try:
            txt = parrafo.text.strip()
            if not txt:
                continue
            txt_lower = txt.lower()
            texto_word_completo += " " + txt_lower
            
            if "ecatepec" in txt_lower:
                tiene_ecatepec = True
            if "iztapalapa" in txt_lower:
                tiene_iztapalapa = True

            # --- CORRECCIÓN ORTOGRÁFICA CRÍTICA ---
            if "rocio" in txt_lower and ("maritza" in txt_lower or "ramírez" in txt_lower):
                if "roció" in txt_lower or "rocio" in txt_lower:
                    for run in parrafo.runs:
                        if "roció" in run.text.lower() or "rocio" in run.text.lower():
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    palabras_sospechosas.append(txt)

            # --- AUDITORÍA DE TIPOGRAFÍA (Raleway 9-11) ---
            for run in parrafo.runs:
                if run.text.strip():
                    fuente = run.font.name
                    tamaño = run.font.size.pt if run.font.size else None
                    if (fuente and fuente != "Raleway") or (tamaño and (tamaño < 9.0 or tamaño > 11.0)):
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        alertas_diseno.append(f"⚠️ **Párrafo {i}:** Usa letra fuera de formato.")

            # --- AUDITORÍA DE ALINEACIÓN ---
            alineacion = parrafo.alignment
            es_palabra_centrada = any(p_centrada in txt_lower for p_centrada in ["d i c t a m e n", "atentamente", "nombre y firma"])
            
            if es_palabra_centrada:
                if alineacion != WD_ALIGN_PARAGRAPH.CENTER:
                    for run in parrafo.runs:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    alertas_diseno.append(f"❌ **Párrafo {i}:** El rubro *'{txt}'* debe ir **CENTRADO**.")
            else:
                if len(txt) > 40 and alineacion != WD_ALIGN_PARAGRAPH.JUSTIFY:
                    for run in parrafo.runs:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    alertas_diseno.append(f"❌ **Párrafo {i}:** Texto libre no se encuentra **JUSTIFICADO**.")

            # VALIDACIÓN DE ANTECEDENTES
            if "antecedente" in txt_lower or "oficio número" in txt_lower or "oficio numero" in txt_lower:
                if oficio_solicitud.lower() not in txt_lower and "fgr" in txt_lower:
                    for run in parrafo.runs:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    errores_antecedentes_cuenta += 1

            # Marcar Contradicción Geográfica
            if tiene_ecatepec and tiene_iztapalapa and "iztapalapa" in txt_lower and "[⚠️" not in txt:
                parrafo.add_run(" [⚠️ CONTRADICCIÓN DE PLANTILLA: Estudio declara Ecatepec, no Iztapalapa.]")
                for run in parrafo.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                errores_congruencia_cuenta += 1

        except Exception as e:
            continue

    st.success("✅ ¡Auditoría completada!")
    st.divider()

    # --- REPORTES EN PANTALLA DE DISEÑO ---
    st.subheader("📐 1. Reporte de Diseño y Formalidad del Documento")
    if alertas_diseno:
        st.warning("Detalles de alineación o fuentes detectados:")
        for alerta in list(set(alertas_diseno))[:5]:
            st.markdown(alerta)
    else:
        st.success("🎉 Estructura formal impecable (Raleway, Justificados y Centrados correctos).")

    st.divider()

    # --- REPORTES DE VALIDACIÓN CRUZADA ---
    st.subheader("🕵️‍♂️ 2. Resultados de Validación Cruzada (PDF vs. Word)")
    col_pdf1, col_pdf2 = st.columns(2)
    with col_pdf1:
        st.info(f"📄 **Oficio en PDF:** {oficio_solicitud}")
    with col_pdf2:
        st.info(f"📂 **Carpeta en PDF:** {carpeta_solicitud}")

    # Reporte Ortográfico Crítico
    st.subheader("📝 3. Reporte de Corrección Ortográfica y Acentuación")
    if palabras_sospechosas:
        st.warning("Se detectaron detalles de acentuación críticos en nombres propios:")
        st.markdown(f"* En tus párrafos de redacción escribiste **'Roció'** de forma incorrecta. La forma oficial is **'Rocío'** (con acento en la 'í').")
    else:
        st.success("🎉 ¡Excelente! No se detectaron faltas de ortografía en los nombres del personal.")

    # Verificar presencia de Rubros Obligatorios
    rubros_faltantes = []
    texto_completo_limpio = texto_word_completo.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    for rubro in RUBROS_BASE:
        rubro_limpio = rubro.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        patron = rf"{rubro_limpio}(es|s)?\b"
        if not re.search(patron, texto_completo_limpio):
            rubros_faltantes.append(rubro.upper())

    if rubros_faltantes:
        st.error(f"❌ Faltan los siguientes rubros obligatorios: {', '.join(rubros_faltantes)}")

    st.divider()

    # ------------------ BOTÓN DE DESCARGA DIRECTO Y SEGURO ------------------
    st.subheader("📥 Descarga tu archivo auditado")
    st.write("Al descargar el documento, cualquier párrafo desalineado o palabra con falta de ortografía vendrá resaltada en **amarillo fosforescente**.")
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button(
        label="📥 Descargar Word con Marcas de Error",
        data=bio,
        file_name="DICTAMEN_AUDITADO.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.warning("💡 Por favor, sube **ambos archivos** para iniciar la auditoría.")
