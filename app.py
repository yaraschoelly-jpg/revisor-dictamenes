import io
import re
import unicodedata
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
import pypdf
import streamlit as st


# Función para quitar acentos
def quitar_acentos(texto):
  if not texto:
    return ""
  texto_norm = unicodedata.normalize("NFD", texto)
  return "".join(c for c in texto_norm if unicodedata.category(c) != "Mn")


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
  with st.spinner(
      "🔍 Analizando consistencia, formalidad y ortografía... Por favor,"
      " espera."
  ):

    # --- 1. LECTURA Y EXTRACCIÓN DEL PDF ---
    texto_pdf = ""
    try:
      lector_pdf = pypdf.PdfReader(archivo_pdf)
      for pag in lector_pdf.pages:
        txt_pag = pag.extract_text()
        if txt_pag:
          texto_pdf += " " + txt_pag
    except Exception as e:
      st.error(f"Error al leer el PDF: {e}")

    match_carpeta = re.search(
        r"(FED|CUI|EXP|CP|CI|CAUSA)[/\-\w\d]+", texto_pdf, re.IGNORECASE
    )
    carpeta_solicitud = (
        match_carpeta.group(0).upper() if match_carpeta else "NO DETECTADO"
    )

    match_oficio = re.search(
        r"(FGR|AIC|PFM|UINP|SUB)[/\-\w\d]+", texto_pdf, re.IGNORECASE
    )
    oficio_solicitud = (
        match_oficio.group(0).upper() if match_oficio else "NO DETECTADO"
    )

    # --- 2. LECTURA Y AUDITORÍA DEL WORD ---
    doc = docx.Document(archivo_docx)
    texto_word_completo = ""
    palabras_sospechosas = []
    alertas_diseno = []

    # A. Revisión del Encabezado
    try:
      for seccion in doc.sections:
        if seccion.header:
          for p in seccion.header.paragraphs:
            p_text = p.text.strip()
            if "carpeta" in p_text.lower():
              if (
                  carpeta_solicitud != "NO DETECTADO"
                  and carpeta_solicitud.lower() not in p_text.lower()
              ):
                p.add_run(
                    f" [⚠️ ERROR: EN SOLICITUD CONSTA {carpeta_solicitud}]"
                )
                for r in p.runs:
                  r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    except Exception:
      pass

    # B. Revisión de Párrafos del Cuerpo
    tiene_ecatepec = False
    tiene_iztapalapa = False

    for i, p in enumerate(doc.paragraphs, start=1):
      txt = p.text.strip()
      if not txt:
        continue

      txt_lower = txt.lower()
      texto_word_completo += " " + txt_lower

      if "ecatepec" in txt_lower:
        tiene_ecatepec = True
      if "iztapalapa" in txt_lower:
        tiene_iztapalapa = True

      # Revisión Ortográfica: "Rocío"
      txt_sin_acentos = quitar_acentos(txt_lower)
      if ("maritza" in txt_sin_acentos or "ramirez" in txt_sin_acentos) and (
          "rocio" in txt_sin_acentos
      ):
        for r in p.runs:
          if "rocio" in quitar_acentos(r.text.lower()):
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        palabras_sospechosas.append(
            f"Párrafo {i}: Verificar acentuación del nombre 'Rocío'."
        )

      # --- CENTRADO OBLIGATORIO DE 'DICTAMEN' Y RUBROS ---
      es_palabra_dictamen = bool(
          re.search(r"\bd\s*i\s*c\s*t\s*a\s*m\s*e\s*n\b", txt_lower)
      )
      es_centrado = es_palabra_dictamen or any(
          kw in txt_lower
          for kw in ["atentamente", "nombre y firma", "dictamen pericial"]
      )

      if es_centrado:
        if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
          p.alignment = WD_ALIGN_PARAGRAPH.CENTER
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_diseno.append(
              f"❌ Párrafo {i}: La palabra *'{txt[:30]}'* debe ir CENTRADA (se"
              " corrigió en el archivo)."
          )
      elif len(txt) > 80:
        if (
            p.alignment is not None
            and p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY
        ):
          for r in p.runs:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_diseno.append(f"❌ Párrafo {i}: Debe ir JUSTIFICADO.")

      # Contradicción Geográfica
      if (
          tiene_ecatepec
          and tiene_iztapalapa
          and "iztapalapa" in txt_lower
          and "[⚠️" not in txt
      ):
        p.add_run(
            " [⚠️ CONTRADICCIÓN DE PLANTILLA: Se detectó Ecatepec e Iztapalapa"
            " en el texto.]"
        )
        for r in p.runs:
          r.font.highlight_color = WD_COLOR_INDEX.YELLOW

  # --- MOSTRAR RESULTADOS ---
  st.success("✅ ¡Auditoría completada!")
  st.divider()

  st.subheader("📐 1.
