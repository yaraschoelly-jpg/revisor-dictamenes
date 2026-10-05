import io
import re
import unicodedata
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
import pypdf
import streamlit as st


def quitar_acentos(texto):
  if not texto:
    return ""
  texto_norm = unicodedata.normalize("NFD", texto)
  return "".join(c for c in texto_norm if unicodedata.category(c) != "Mn")


st.set_page_config(
    page_title="Auditor Pericial Integral", page_icon="⚖️", layout="centered"
)

st.title("⚖️ Auditor Pericial de Formalidad y Cotejo")
st.write(
    "Sube el **PDF de Solicitud** y el **Word del Dictamen** para auditar"
    " redacción, formato y consistencia de datos."
)

st.subheader("1. Carga de Documentos Oficiales")
col_pdf, col_docx = st.columns(2)

with col_pdf:
  archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
  archivo_docx = st.file_uploader(
      "Subir Dictamen Pericial en Word (.docx)", type=["docx"]
  )

if archivo_pdf is not None and archivo_docx is not None:
  with st.spinner(
      "Procesando y cotejando información... Por favor, espera."
  ):

    # --- 1. EXTRACCIÓN DE DATOS EN PDF (SOLICITUD) ---
    texto_pdf = ""
    try:
      lector_pdf = pypdf.PdfReader(archivo_pdf)
      for pag in lector_pdf.pages:
        txt_pag = pag.extract_text()
        if txt_pag:
          texto_pdf += " " + txt_pag
    except Exception as e:
      st.error(f"Error al leer el PDF: {e}")

    # Extracción de Carpeta de Investigación
    match_carpeta = re.search(
        r"(FED|CUI|EXP|CP|CI|CAUSA)[/\-\w\d]+", texto_pdf, re.IGNORECASE
    )
    carpeta_solicitud = (
        match_carpeta.group(0).upper() if match_carpeta else "NO DETECTADO"
    )

    # Extracción de Número de Oficio
    match_oficio = re.search(
        r"(FGR|AIC|PFM|UINP|SUB|OFICIO)[/\-\w\d]+", texto_pdf, re.IGNORECASE
    )
    oficio_solicitud = (
        match_oficio.group(0).upper() if match_oficio else "NO DETECTADO"
    )

    # Extracción del Remitente (Nombre de quien envía)
    match_remitente = re.search(
        r"(LIC\.|MTRO\.|MTRA\.|DR\.|DRA\.|LICENCIADO|LICENCIADA)\s+([A-ZÁÉÍÓÚÑ\s]+)",
        texto_pdf,
        re.IGNORECASE,
    )
    remitente_solicitud = (
        match_remitente.group(0).strip() if match_remitente else "NO DETECTADO"
    )

    # --- 2. AUDITORÍA EN WORD (DICTAMEN) ---
    doc = docx.Document(archivo_docx)
    texto_word_completo = ""
    alertas_alineacion = []
    observaciones_cotejo = []

    # A. Cotejo en Encabezados y Secciones del Word
    try:
      for seccion in doc.sections:
        if seccion.header:
          for p in seccion.header.paragraphs:
            p_text = p.text.strip()
            if "carpeta" in p_text.lower() and carpeta_solicitud != "NO DETECTADO":
              if carpeta_solicitud.lower() not in p_text.lower():
                p.add_run(
                    f" [⚠️ ERROR COTEJO: En solicitud consta"
                    f" {carpeta_solicitud}]"
                )
                for r in p.runs:
                  r.font.highlight_color = WD_COLOR_INDEX.YELLOW
                observaciones_cotejo.append(
                    f"Carpeta de investigación no coincide con la solicitud"
                    f" ({carpeta_solicitud})."
                )
    except Exception:
      pass

    # B. Revisión Párrafo por Párrafo
    for i, p in enumerate(doc.paragraphs, start=1):
      txt = p.text.strip()
      if not txt:
        continue

      txt_lower = txt.lower()
      txt_limpio = quitar_acentos(txt_lower)
      texto_word_completo += " " + txt_lower

      # --- REGLA DE ALINEACIÓN ---
      # Excepciones que DEBEN ir CENTRADAS
      es_palabra_dictamen = "dictamen" in txt_limpio.replace(" ", "") and len(txt) < 30
      es_atentamente = "atentamente" in txt_limpio and len(txt) < 30
      es_perito = "perito en criminalistica" in txt_limpio
      
      # Si contiene el bloque de firma o nombre propio en áreas de cierre
      es_bloque_firma = es_palabra_dictamen or es_atentamente or es_perito

      if es_bloque_firma:
        if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
          p.alignment = WD_ALIGN_PARAGRAPH.CENTER
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_alineacion.append(
              f"Párrafo {i}: '{txt[:35]}...' debe estar CENTRADO (se corrigió"
              " en el archivo)."
          )
      else:
        # Cualquier otro párrafo debe ser JUSTIFICADO
        if len(txt) > 40 and p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
          p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_alineacion.append(
              f"Párrafo {i}: Párrafo no justificado (se aplicó alineación"
              " JUSTIFICADA)."
          )

      # --- COTEJO DE REMITENTE Y OFICIO EN EL CUERPO ---
      if (
          remitente_solicitud != "NO DETECTADO"
          and "atencion" in txt_limpio
          or "solicitante" in txt_limpio
      ):
        nombre_remitente_limpio = quitar_acentos(
            remitente_solicitud.split()[-1].lower()
        )
        if (
            len(nombre_remitente_limpio) > 3
            and nombre_remitente_limpio not in txt_limpio
        ):
          p.add_run(
              f" [⚠️ REVISAR REMITENTE: En solicitud consta"
              f" {remitente_solicitud}]"
          )
          for r in r_p in p.runs:
            r_p.font.highlight_color = WD_COLOR_INDEX.YELLOW

  # --- MOSTRAR RESULTADOS ---
  st.success("🎉 ¡Auditoría completada!")
  st.divider()

  # 1. COTEJO DE INFORMACIÓN CRUZADA
  st.subheader("🕵️‍♂️ 1. Cotejo de Datos (PDF Solicitud vs. Word Dictamen)")
  col1, col2, col3 = st.columns(3)
  col1.metric("Carpeta de Inv.", carpeta_solicitud)
  col2.metric("Oficio Solicitud", oficio_solicitud)
  col3.metric("Remitente (Envía)", remitente_solicitud)

  if observaciones_cotejo:
    for obs in observaciones_cotejo:
      st.error(f"❌ {obs}")
  else:
    st.info("Los datos clave del PDF fueron comparados con el documento Word.")

  st.divider()

  # 2. ALINEACIÓN Y FORMATO
  st.subheader("📐 2. Reporte de Formato y Alineaciones")
  if alertas_alineacion:
    for al in list(set(alertas_alineacion))[:6]:
      st.write(f"* {al}")
  else:
    st.success("Alineaciones correctas (Justificados y Centrados requeridos cumplidos).")

  st.divider()

  # 3. DESCARGA
  st.subheader("📥 3. Descargar Word Auditado")
  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)

  st.download_button(
      label="📥 Descargar Word con Correcciones y Marcas",
      data=bio,
      file_name="DICTAMEN_AUDITADO.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
  )
else:
  st.warning("💡 Por favor, sube ambos archivos para iniciar la auditoría.")
