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
    " encabezados, redacción, formato y consistencia de datos."
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
        r"(FED|CUI|EXP|CP|CI|CAUSA|AP)[/\-\w\d\.]+", texto_pdf, re.IGNORECASE
    )
    carpeta_solicitud = (
        match_carpeta.group(0).upper() if match_carpeta else "NO DETECTADO"
    )

    # Extracción de Número de Folio
    match_folio = re.search(
        r"(FOLIO|FOLIO\s*NÚMERO|FOLIO\s*NO\.?)\s*[:\.\-]?\s*([A-Z0-9/\-]+)",
        texto_pdf,
        re.IGNORECASE,
    )
    if not match_folio:
      # Búsqueda alternativa para secuencias numéricas de folio de 5 a 8 dígitos
      match_folio = re.search(r"\b\d{5,8}\b", texto_pdf)
      folio_solicitud = match_folio.group(0) if match_folio else "NO DETECTADO"
    else:
      folio_solicitud = match_folio.group(2).upper()

    # Extracción de Número de Oficio
    match_oficio = re.search(
        r"(FGR|AIC|PFM|UINP|SUB|OFICIO)[/\-\w\d]+", texto_pdf, re.IGNORECASE
    )
    oficio_solicitud = (
        match_oficio.group(0).upper() if match_oficio else "NO DETECTADO"
    )

    # Extracción del Remitente
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

    # A. Auditando el Encabezado de la Sección (Header)
    try:
      for seccion in doc.sections:
        if seccion.header:
          header_text = ""
          paragraphs_header = seccion.header.paragraphs
          for p in paragraphs_header:
            header_text += " " + p.text.strip().lower()

          # Validar Carpeta de Investigación en Encabezado
          if carpeta_solicitud != "NO DETECTADO":
            carpeta_clean = re.sub(r"[^\w]", "", carpeta_solicitud.lower())
            header_clean = re.sub(r"[^\w]", "", header_text)
            if carpeta_clean not in header_clean:
              if paragraphs_header:
                p_target = paragraphs_header[0]
                p_target.add_run(
                    f" [⚠️ ERROR EN ENCABEZADO: Falta o difiere Carpeta"
                    f" ({carpeta_solicitud})]"
                )
                for r in p_target.runs:
                  r.font.highlight_color = WD_COLOR_INDEX.YELLOW
              observaciones_cotejo.append(
                  f"Carpeta de investigación en encabezado difiere o no consta"
                  f" ({carpeta_solicitud})."
              )

          # Validar Número de Folio en Encabezado
          if folio_solicitud != "NO DETECTADO":
            folio_clean = re.sub(r"[^\w]", "", folio_solicitud.lower())
            header_clean = re.sub(r"[^\w]", "", header_text)
            if folio_clean not in header_clean:
              if paragraphs_header:
                p_target = paragraphs_header[-1]
                p_target.add_run(
                    f" [⚠️ ERROR EN ENCABEZADO: Falta o difiere Número de Folio"
                    f" ({folio_solicitud})]"
                )
                for r in p_target.runs:
                  r.font.highlight_color = WD_COLOR_INDEX.YELLOW
              observaciones_cotejo.append(
                  f"Número de folio en encabezado difiere o no consta"
                  f" ({folio_solicitud})."
              )
    except Exception as e:
      pass

    # B. Revisión Párrafo por Párrafo del Cuerpo
    for i, p in enumerate(doc.paragraphs, start=1):
      txt = p.text.strip()
      if not txt:
        continue

      txt_lower = txt.lower()
      txt_limpio = quitar_acentos(txt_lower)
      texto_word_completo += " " + txt_lower

      # --- REGLAS DE ALINEACIÓN ---

      # 1. ELEMENTOS A LA DERECHA
      es_asunto = "asunto:" in txt_limpio or "se emite dictamen" in txt_limpio
      es_leyenda_oficial = "margarita maza parada" in txt_limpio or "ano de" in txt_limpio
      es_fecha = bool(re.search(r'\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b', txt_limpio)) and len(txt) < 60

      es_derecha = es_asunto or es_leyenda_oficial or es_fecha

      # 2. ELEMENTOS AL CENTRO
      es_palabra_dictamen = "dictamen" in txt_limpio.replace(" ", "") and len(txt) < 30
      es_atentamente = "atentamente" in txt_limpio and len(txt) < 30
      es_perito = "perito en criminalistica" in txt_limpio

      es_centrado = es_palabra_dictamen or es_atentamente or es_perito

      # EVALUACIÓN Y CORRECCIÓN DE ALINEACIONES
      if es_derecha:
        if p.alignment != WD_ALIGN_PARAGRAPH.RIGHT:
          p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_alineacion.append(
              f"Párrafo {i}: '{txt[:35]}...' alineado A LA DERECHA."
          )

      elif es_centrado:
        if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
          p.alignment = WD_ALIGN_PARAGRAPH.CENTER
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_alineacion.append(
              f"Párrafo {i}: '{txt[:35]}...' alineado AL CENTRO."
          )

      else:
        # CUALQUIER OTRO PÁRRAFO LIBRE DEBE SER JUSTIFICADO
        if len(txt) > 40 and p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
          p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_alineacion.append(
              f"Párrafo {i}: Alineación ajustada a JUSTIFICADO."
          )

  # --- MOSTRAR RESULTADOS ---
  st.success("🎉 ¡Auditoría completada!")
  st.divider()

  # 1. COTEJO DE INFORMACIÓN CRUZADA
  st.subheader("🕵️‍♂️ 1. Datos Extraídos del PDF de Solicitud")
  col1, col2, col3, col4 = st.columns(4)
  col1.metric("Carpeta Inv.", carpeta_solicitud)
  col2.metric("Folio", folio_solicitud)
  col3.metric("Oficio", oficio_solicitud)
  col4.metric("Remitente", remitente_solicitud)

  if observaciones_cotejo:
    st.subheader("⚠️ Observaciones de Encabezado y Cotejo")
    for obs in set(observaciones_cotejo):
      st.error(f"❌ {obs}")
  else:
    st.info("Los datos del encabezado coinciden con la solicitud en PDF.")

  st.divider()

  # 2. ALINEACIÓN Y FORMATO
  st.subheader("📐 2. Reporte de Formato y Alineaciones")
  if alertas_alineacion:
    for al in list(set(alertas_alineacion))[:8]:
      st.write(f"* {al}")
  else:
    st.success("Alineaciones correctas (Derecha, Centro y Justificado).")

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
