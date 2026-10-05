import io
import re
import unicodedata
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pdfplumber
from pdf2image import convert_from_bytes
import pytesseract
import streamlit as st


def quitar_acentos(texto):
  if not texto:
    return ""
  texto_norm = unicodedata.normalize("NFD", texto)
  return "".join(c for c in texto_norm if unicodedata.category(c) != "Mn")


def activar_control_de_cambios(doc):
  """Activa el control de cambios (Track Changes) en el archivo Word."""
  settings = doc.settings._element
  track_revisions = settings.find(qn("w:trackRevisions"))
  if track_revisions is None:
    track_revisions = OxmlElement("w:trackRevisions")
    settings.append(track_revisions)


def corregir_y_resaltar_ortografia(p, idx, alertas_ortografia):
  """Detecta errores de sintaxis u ortografía técnica, resalta la palabra original

  en amarillo y sugiere la corrección.
  """
  correcciones = {
      r"\bcaracteristicas\b": "características",
      r"\bfisica\b": "física",
      r"\bfisicas\b": "físicas",
      r"\bubicacion\b": "ubicación",
      r"\bdescripcion\b": "descripción",
      r"\bobservacion\b": "observación",
      r"\bobservaciones\b": "observaciones",
      r"\bconsideracion\b": "consideración",
      r"\bconsideraciones\b": "consideraciones",
      r"\bconclusion\b": "conclusión",
      r"\bconclusiones\b": "conclusiones",
      r"\bindicio\b": "indicio",
      r"\bindicios\b": "indicios",
      r"\btecnica\b": "técnica",
      r"\btecnico\b": "técnico",
  }

  texto_original = p.text
  hubo_cambio = False

  for patron, reemplazo in correcciones.items():
    if re.search(patron, texto_original, re.IGNORECASE):
      hubo_cambio = True
      # Resaltar párrafos con errores de sintaxis/ortografía técnica
      for run in p.runs:
        if re.search(patron, run.text, re.IGNORECASE):
          run.font.highlight_color = WD_COLOR_INDEX.YELLOW

      # Sustitución respetando mayúsculas y minúsculas
      texto_original = re.sub(
          patron, reemplazo, texto_original, flags=re.IGNORECASE
      )

  if hubo_cambio:
    p.text = texto_original
    # Reaplicar el resaltado amarillo al texto corregido para visibilidad
    for run in p.runs:
      for _, reemplazo in correcciones.items():
        if reemplazo.lower() in run.text.lower():
          run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    alertas_ortografia.append(
        f"Párrafo {idx}: Se identificaron y subrayaron errores de acentuación"
        " o sintaxis."
    )


st.set_page_config(
    page_title="Auditor Pericial Integral", page_icon="⚖️", layout="centered"
)

st.title("⚖️ Auditor Pericial de Formalidad y Sintaxis")
st.write(
    "Sube el **PDF de Solicitud** y el **Word del Dictamen** para activar el"
    " **Control de Cambios**, resaltar errores gramaticales y auditar la"
    " estructura."
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
      "Procesando, activando Control de Cambios y resaltando errores... Por"
      " favor, espera."
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

    # --- 2. AUDITORÍA, CONTROL DE CAMBIOS Y RESALTADO EN WORD ---
    doc = docx.Document(archivo_docx)

    # ACTIVAR CONTROL DE CAMBIOS EN EL DOCUMENTO
    activar_control_de_cambios(doc)

    texto_word_completo = ""
    alertas_alineacion = []
    alertas_ortografia = []
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
