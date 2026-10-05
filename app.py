import io
import re
import unicodedata
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
import pdfplumber
from pdf2image import convert_from_bytes
import pytesseract
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
    " encabezados, ubicación de destinatario, redacción y formato."
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
      "Procesando y realizando OCR en documentos... Por favor, espera."
  ):

    # --- 1. LECTURA DIGITAL + MOTOR OCR ---
    bytes_pdf = archivo_pdf.read()
    texto_pdf = ""

    try:
      with pdfplumber.open(io.BytesIO(bytes_pdf)) as pdf:
        for pagina in pdf.pages:
          t = pagina.extract_text(layout=True)
          if t:
            texto_pdf += "\n" + t
    except Exception:
      pass

    if not texto_pdf.strip():
      try:
        imagenes = convert_from_bytes(bytes_pdf)
        for img in imagenes:
          texto_pdf += "\n" + pytesseract.image_to_string(img, lang="spa")
      except Exception as e:
        st.error(f"Error al ejecutar OCR en el PDF: {e}")

    # Extracción de Datos en PDF
    match_carpeta = re.search(
        r"(carpeta|expediente|causa|cui|ap|ci)\s*[\w\d\.\-/:]+",
        texto_pdf,
        re.IGNORECASE,
    )
    if not match_carpeta:
      match_carpeta = re.search(
          r"[A-Z0-9]{2,8}/[A-Z0-9/\-_]{4,25}", texto_pdf
      )
    carpeta_solicitud = (
        match_carpeta.group(0).upper().strip() if match_carpeta else "NO DETECTADO"
    )

    match_folio = re.search(
        r"(folio)\s*[\w\d\.\-:]+", texto_pdf, re.IGNORECASE
    )
    if not match_folio:
      match_folio = re.search(r"\b\d{5,8}\b", texto_pdf)
    folio_solicitud = (
        match_folio.group(0).upper().strip() if match_folio else "NO DETECTADO"
    )

    match_oficio = re.search(
        r"(oficio|fgr|aic|pfm|uinp|sub)\s*[\w\d\.\-/:]+",
        texto_pdf,
        re.IGNORECASE,
    )
    oficio_solicitud = (
        match_oficio.group(0).upper().strip() if match_oficio else "NO DETECTADO"
    )

    match_remitente = re.search(
        r"(lic\.|mtro\.|mtra\.|dr\.|dra\.|licenciado|licenciada|c\.)\s+([a-záéíóúñ\s]+)",
        texto_pdf,
        re.IGNORECASE,
    )
    remitente_solicitud = (
        match_remitente.group(0).strip().upper()
        if match_remitente
        else "NO DETECTADO"
    )

    # --- 2. AUDITORÍA Y CORRECCIÓN EN WORD ---
    doc = docx.Document(archivo_docx)
    texto_word_completo = ""
    alertas_alineacion = []
    observaciones_cotejo = []

    # A. Auditando Encabezado
    try:
      for seccion in doc.sections:
        if seccion.header:
          header_text = ""
          paragraphs_header = seccion.header.paragraphs
          for p in paragraphs_header:
            header_text += " " + p.text.strip().lower()

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
                  f"Carpeta de investigación difiere o no consta ({carpeta_solicitud})."
              )

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
                  f"Número de folio difiere o no consta ({folio_solicitud})."
              )
    except Exception:
      pass

    # B. Eliminación de espacios en blanco antes de "PRESENTE"
    parrafos = doc.paragraphs
    i = 0
    while i < len(parrafos):
      txt_p = parrafos[i].text.strip()
      txt_limpio_p = quitar_acentos(txt_p.lower()).replace(" ", "")

      if txt_limpio_p == "presente":
        # Revisar párrafos anteriores en blanco y eliminarlos
        j = i - 1
        while j >= 0 and not parrafos[j].text.strip():
          p_element = parrafos[j]._element
          p_element.getparent().remove(p_element)
          alertas_alineacion.append(
              f"Se eliminó un espacio en blanco innecesario antes de la palabra 'PRESENTE'."
          )
          j -= 1
      i += 1

    # C. Revisión Párrafo por Párrafo del Cuerpo
    for idx, p in enumerate(doc.paragraphs, start=1):
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

      # 3. REGLA ESPECÍFICA "PRESENTE" (Izquierda / Debajo del Destinatario)
      es_presente = txt_limpio.replace(" ", "") == "presente"

      if es_presente:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        alertas_alineacion.append(
            f"Párrafo {idx}: La palabra 'PRESENTE' se alineó correctamente debajo de la autoridad solicitante sin espacios intermedios."
        )

      elif es_derecha:
        if p.alignment != WD_ALIGN_PARAGRAPH.RIGHT:
          p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_alineacion.append(
              f"Párrafo {idx}: '{txt[:35]}...' alineado A LA DERECHA."
          )

      elif es_centrado:
        if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
          p.alignment = WD_ALIGN_PARAGRAPH.CENTER
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_alineacion.append(
              f"Párrafo {idx}: '{txt[:35]}...' alineado AL CENTRO."
          )

      else:
        if len(txt) > 40 and p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
          p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_alineacion.append(
              f"Párrafo {idx}: Alineación ajustada a JUSTIFICADO."
          )

  # --- MOSTRAR RESULTADOS ---
  st.success("🎉 ¡Auditoría completada!")
  st.divider()

  st.subheader("🕵️‍♂️ 1. Datos Extraídos del PDF (Digital/OCR)")
  col1, col2, col3, col4 = st.columns(4)
  col1.metric("Carpeta Inv.", carpeta_solicitud)
  col2.metric("Folio", folio_solicitud)
  col3.metric("Oficio", oficio_solicitud)
  col4.metric("Remitente", remitente_solicitud)

  with st.expander("🔍 Ver texto extraído por OCR del PDF (Verificación)"):
    if texto_pdf.strip():
      st.text(texto_pdf)
    else:
      st.warning("No se pudo extraer texto ni con OCR. Verifica la imagen.")

  if observaciones_cotejo:
    st.subheader("⚠️️ Observaciones de Encabezado y Cotejo")
    for obs in set(observaciones_cotejo):
      st.error(f"❌ {obs}")

  st.divider()

  st.subheader("📐 2. Reporte de Formato y Alineaciones")
  if alertas_alineacion:
    for al in list(set(alertas_alineacion))[:8]:
      st.write(f"* {al}")
  else:
    st.success("Alineaciones correctas (Derecha, Centro, Presente y Justificado).")

  st.divider()

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
