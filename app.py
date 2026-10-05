import io
import re
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
import pypdf
import streamlit as st

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

  # --- 2. EXTRAER DATOS CLAVE DEL PDF (Regex insensible a mayúsculas/minúsculas) ---
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

  # B. Revisión del Cuerpo
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

      # --- REVISIÓN ORTOGRÁFICA / NOMBRES ---
      if ("maritza" in txt_lower or "ramírez" in txt_lower) and (
          "rocio" in txt_lower or "roció" in txt_lower
      ):
        for run in parrafo.runs:
          if any(w in run.text.lower() for w in ["rocio", "roció"]):
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        palabras_sospechosas.append(
            f"Párrafo {i}: Se identificó 'Roció/Rocio' (debe evaluarse acentuación 'Rocío')."
        )

      # --- AUDITORÍA DE TIPOGRAFÍA (Aplica resaltado directo en runs) ---
      for run in parrafo.runs:
        if run.text.strip():
          fuente = run.font.name
          tamaño = run.font.size.pt if run.font.size else None

          if (fuente and fuente != "Raleway") or (
              tamaño and (tamaño < 9.0 or tamaño > 11.0)
          ):
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            alertas_diseno.append(
                f"⚠️ **Párrafo {i}:** Fuente o tamaño fuera de formato."
            )

      # --- AUDITORÍA DE ALINEACIÓN ---
      alineacion = parrafo.alignment
      es_palabra_centrada = any(
          p in txt_lower
          for p in [
              "d i c t a m e n",
              "atentamente",
              "nombre y firma",
              "dictamen pericial",
          ]
      )

      if es_palabra_centrada:
        if alineacion is not None and alineacion != WD_ALIGN_PARAGRAPH.CENTER:
          for run in parrafo.runs:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_diseno.append(
              f"❌ **Párrafo {i}:** El rubro *'{txt[:30]}...'* debe ir CENTRADO."
          )
      else:
        if (
            len(txt) > 60
            and alineacion is not None
            and alineacion != WD_ALIGN_PARAGRAPH.JUSTIFY
        ):
          for run in parrafo.runs:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_diseno.append(
              f"❌ **Párrafo {i}:** Párrafo no se encuentra JUSTIFICADO."
          )

      # --- CONTRADICCIÓN DE PLANTILLA ---
      if (
          tiene_ecatepec
          and tiene_iztapalapa
          and "iztapalapa" in txt_lower
          and "[⚠️" not in txt
      ):
        run_warn = parrafo.add_run(
            " [⚠️ CONTRADICCIÓN DE PLANTILLA: Se detectó Ecatepec e Iztapalapa"
            " en el cuerpo.]"
        )
        run_warn.font.highlight_color = WD_COLOR_INDEX.YELLOW
        for run in parrafo.runs:
          run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    except Exception as e:
      continue

  st.success("✅ ¡Auditoría completada!")
  st.divider()

  # --- REPORTES EN PANTALLA ---
  st.subheader("📐 1. Reporte de Diseño y Formalidad")
  if alertas_diseno:
    st.warning("Detalles de alineación o fuentes detectados:")
    for alerta in list(set(alertas_diseno))[:5]:
      st.markdown(alerta)
  else:
    st.success("Estructura formal sin observaciones detectadas.")

  st.divider()

  st.subheader("🕵️‍♂️ 2. Validación Cruzada (PDF vs. Word)")
  col_pdf1, col_pdf2 = st.columns(2)
  with col_pdf1:
    st.info(f"📄 **Oficio en PDF:** {oficio_solicitud}")
  with col_pdf2:
    st.info(f"📂 **Carpeta en PDF:** {carpeta_solicitud}")

  # Reporte Ortográfico
  st.subheader("📝 3. Reporte Ortográfico")
  if palabras_sospechosas:
    st.warning("Puntos ortográficos detectados:")
    for obs in set(palabras_sospechosas):
      st.markdown(f"* {obs}")
  else:
    st.success("Sin faltas ortográficas críticas en nombres propios.")

  # Rubros faltantes
  rubros_faltantes = []
  texto_completo_limpio = (
      texto_word_completo.replace("á", "a")
      .replace("é", "e")
      .replace("í", "i")
      .replace("ó", "o")
      .replace("ú", "u")
  )
  for rubro in RUBROS_BASE:
    rubro_limpio = (
        rubro.replace("á", "a")
        .replace("é", "e")
        .replace("í", "
