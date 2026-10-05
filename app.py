import io
import re
import unicodedata
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
import pypdf
import streamlit as st


# Función auxiliar para quitar acentos de forma segura
def quitar_acentos(texto):
  texto_normalizado = unicodedata.normalize("NFD", texto)
  return "".join(c for c in texto_normalizado if unicodedata.category(c) != "Mn")


# Configuración de la interfaz web
st.set_page_config(
    page_title="Auditor Pericial Integral", page_icon="⚖️", layout="centered"
)

st.title("⚖️ Auditor Pericial de Formalidad Estructural")
st.write(
    "Sube el **PDF de Solicitud** y tu **Word del Dictamen**. El sistema"
    " validará la estructura formal y aplicará las marcas directamente en tu"
    " documento de Word."
)

RUBROS_BASE = [
    "planteamiento del problema",
    "antecedente",
    "estudio de campo",
    "direccion",
    "descripcion del lugar",
    "observacion",
    "consideracion",
    "conclusion",
]

st.subheader("📁 1. Carga de Documentos Oficiales")
col_pdf, col_docx = st.columns(2)

with col_pdf:
  archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
  archivo_docx = st.file_uploader(
      "Subir Dictamen Pericial en Word (.docx)", type=["docx"]
  )

if archivo_pdf is not None and archivo_docx is not None:
  st.info(
      "🔍 Analizando consistencia, formalidad y ortografía... Por favor,"
      " espera."
  )

  # --- 1. EXTRACCIÓN DE TEXTO DEL PDF ---
  texto_pdf = ""
  try:
    lector_pdf = pypdf.PdfReader(archivo_pdf)
    for pagina in lector_pdf.pages:
      t = pagina.extract_text()
      if t:
        texto_pdf += " " + t
  except Exception as e:
    st.error(f"Error al leer el archivo PDF: {e}")

  # --- 2. EXTRAER DATOS CLAVE DEL PDF ---
  match_carpeta_pdf = re.search(
      r"(FED|CUI|EXP|CP|CI|CAUSAPENAL)[/\-\w]+", texto_pdf, re.IGNORECASE
  )
  carpeta_solicitud = (
      match_carpeta_pdf.group(0).upper() if match_carpeta_pdf else "NO DETECTADO"
  )

  match_oficio_pdf = re.search(
      r"(FGR|AIC|PFM|UINP)[/\-\w]+", texto_pdf, re.IGNORECASE
  )
  oficio_solicitud = (
      match_oficio_pdf.group(0).upper() if match_oficio_pdf else "NO DETECTADO"
  )

  # --- 3. EXTRACCIÓN Y AUDITORÍA EN EL WORD ---
  doc = docx.Document(archivo_docx)
  texto_word_completo = ""

  palabras_sospechosas = []
  alertas_diseno = []

  # A. Revisión de Encabezados (Header)
  try:
    for seccion in doc.sections:
      header = seccion.header
      if header:
        for parrafo in header.paragraphs:
          texto_linea = parrafo.text.strip()
          if not texto_linea:
            continue
          texto_linea_lower = texto_linea.lower()

          if "carpeta" in texto_linea_lower:
            if (
                carpeta_solicitud != "NO DETECTADO"
                and carpeta_solicitud.lower() not in texto_linea_lower
            ):
              run_error = parrafo.add_run(
                  f" [⚠️ ERROR DE CONTROL: EN SOLICITUD CONSTA"
                  f" {carpeta_solicitud}]"
              )
              run_error.font.highlight_color = WD_COLOR_INDEX.YELLOW
              for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
  except Exception as e:
    pass
