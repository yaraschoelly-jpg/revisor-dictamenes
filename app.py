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
    "Sube el **PDF de Solicitud** y tu **Word del Dictamen**. El sistema validará la estructura formal y aplicará las marcas directamente en tu documento de Word."
)

# Lista de rubros obligatorios según el manual institucional
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

# Casillas de doble carga oficiales
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
    texto_pdf = ""

  texto_pdf_lower = texto_pdf.lower()

  # --- 2. EXTRAER DATOS CLAVE DEL PDF (Regex Flexible) ---
  match_carpeta_pdf = re.search(
      r"(fed|cui|exp|cp)/[a-z0-9/_\-]+", texto_pdf_lower
  )
  carpeta_solicitud = (
      match_carpeta_pdf.group(0).upper() if match_carpeta_pdf else "NO DETECTADO"
  )

  match_oficio_pdf = re.search(r"fgr-[a-z0-9\-]+", texto_pdf_lower)
  oficio_solicitud = (
      match_oficio_pdf.group(0).upper() if match_oficio_pdf else "NO DETECTADO"
  )

  # --- 3. EXTRACCIÓN Y AUDITORÍA EN EL WORD ---
  doc = docx.Document(archivo_docx)
  texto_word_completo = ""

  errores_encabezado = 0
  errores_antecedentes = 0
  errores_congruencia = 0

  palabras_sospechosas = []
  alertas_diseno = []

  # A. Revisión del Encabezado (Header)
  try:
    for seccion in doc.sections:
      header = seccion.header
      if header:
        for parrafo in header.paragraphs:
          texto_linea = parrafo.text.strip()
          if not texto_linea:
            continue
          texto_linea_lower = texto_linea.lower()

          if any(
              t in texto_linea_lower
              for t in [
                  "agencia",
                  "centro federal",
                  "unidad de",
                  "especialidad",
              ]
          ):
            continue

          if (
              "carpeta" in texto_linea_lower
              and carpeta_solicitud != "NO DETECTADO"
          ):
            # Normalizar caracteres para comparación limpia
            cap_limpia = re.sub(r"[^\w]", "", carpeta_solicitud.lower())
            linea_limpia = re.sub(r"[^\w]", "", texto_linea_lower)

            if cap_limpia not in linea_limpia:
              parrafo.add_run(
                  f" [⚠️ ERROR DE CONTROL: EN SOLICITUD CONSTA"
                  f" {carpeta_solicitud}]"
              )
              for run in parrafo.runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
              errores_encabezado += 1
  except Exception as e:
    st.warning(f"Aviso en revisión de encabezados: {e}")

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

      # --- CORRECCIÓN ORTOGRÁFICA Y NOMBRES PROPIOS ---
      # Detectar "Rocio" o "Roció" cuando debería ser el nombre propio "Rocío"
      if ("maritza" in txt_lower or "ramírez" in txt_lower) and (
          "rocio" in txt_lower or "roció" in txt_lower
      ):
        for run in parrafo.runs:
          if any(w in run.text.lower() for w in ["rocio", "roció"]):
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        palabras_sospechosas.append(
            f"Párrafo {i}: Verifique el acento en el nombre propio 'Rocío'."
        )

      # --- AUDITORÍA DE TIPOGRAFÍA (Raleway 9-11) ---
      for run in parrafo.runs:
        if run.text.strip():
          fuente = run.font.name
          tamaño = run.font.size.pt if run.font.size else None

          # Evaluar únicamente cuando el valor esté definido explícitamente en el run
          if (fuente and fuente != "Raleway") or (
              tamaño and (tamaño < 9.0 or tamaño > 11.0)
          ):
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            alertas_diseno.append(
                f"⚠️ **Párrafo {i}:** Tipografía o tamaño fuera del formato"
                f" estándar (Detectado: {fuente or 'Heredado'},"
                f" {tamaño or 'Heredado'}pt)."
            )

      # --- AUDITORÍA DE ALINEACIÓN ---
      alineacion = parrafo.alignment
      es_palabra_centrada = any(
          p_centrada in txt_lower
          for p_centrada in [
              "d i c t a m e n",
              "atentamente",
              "nombre y firma",
              "dictamen pericial",
          ]
      )

      if es_palabra_centrada:
        if (
            alineacion is not None
            and alineacion != WD_ALIGN_PARAGRAPH.CENTER
        ):
          for run in parrafo.runs:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
          alertas_diseno.append(
              f"❌ **Párrafo {i}:** El rubro *'{txt[:30]}...'* debe ir"
              " **CENTRADO**."
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
              f"❌ **Párrafo {i}:** Texto libre no se encuentra **JUSTIFICADO**."
          )

      # --- VALIDACIÓN DE ANTECEDENTES ---
      if (
          "antecedente" in txt_lower or "oficio" in txt_lower
      ) and oficio_solicitud != "NO DETECTADO":
        oficio_limpio = re.sub(r"[^\w]", "", oficio_solicitud.lower())
        txt_limpio = re.sub(r"[^\w]", "", txt_lower)
        if "fgr" in txt_lower and oficio_limpio not in txt_limpio:
          for run in parrafo.runs:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
          errores_antecedentes += 1

      # --- CONTRADICCIÓN GEOGRÁFICA DE PLANTILLA ---
      if (
          tiene_ecatepec
          and tiene_iztapalapa
          and "iztapalapa" in txt_lower
          and "[⚠️" not in txt
      ):
        parrafo.add_run(
            " [⚠️ CONTRADICCIÓN DE PLANTILLA: Se menciona Ecatepec e Iztapalapa"
            " en el mismo dictamen.]"
        )
        for run in parrafo.runs:
          run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        errores_congruencia += 1

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
    st.success(
        "🎉 Estructura formal impecable (Raleway, Justificados y Centrados"
        " correctos)."
    )

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
    st.warning("Se detectaron detalles de acentuación o nombres propios:")
    for obs in set(palabras_sospechosas):
      st.markdown(f"* {obs}")
  else:
    st.success(
        "🎉 ¡Excelente! No se detectaron faltas ortográficas críticas en"
        " nombres propios."
    )

  # Verificar presencia de Rubros Obligatorios
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
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
    )
    patron = rf"\b{rubro_limpio}(es|s)?\b"
    if not re.search(patron, texto_completo_limpio):
      rubros_faltantes.append(rubro.upper())

  if rubros_faltantes:
    st.error(
        f"❌ Faltan los siguientes rubros obligatorios: {', '.join(rubros_faltantes)}"
    )

  st.divider()

  # ------------------ BOTÓN DE DESCARGA DIRECTO Y SEGURO ------------------
  st.subheader("📥 Descarga tu archivo auditado")
  st.write(
      "Al descargar el documento, las observaciones y discrepancias"
      " encontradas aparecerán resaltadas en **amarillo**."
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)

  st.download_button(
      label="📥 Descargar Word con Marcas de Error",
      data=bio,
      file_name="DICTAMEN_AUDITADO.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
  )
else:
  st.warning("💡 Por favor, sube **ambos archivos** para iniciar la auditoría.")
