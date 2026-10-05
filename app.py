import io
import re
import unicodedata
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pdfplumber
import streamlit as st

try:
    from pdf2image import convert_from_bytes
    import pytesseract
    HAS_OCR = True
except Exception:
    HAS_OCR = False


def quitar_acentos(texto):
    if not texto:
        return ""
    texto_norm = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto_norm if unicodedata.category(c) != "Mn")


def activar_control_de_cambios(doc):
    try:
        settings = doc.settings._element
        track_revisions = settings.find(qn("w:trackRevisions"))
        if track_revisions is None:
            track_revisions = OxmlElement("w:trackRevisions")
            settings.append(track_revisions)
    except Exception as e:
        st.warning(f"No se pudo activar el control de cambios automaticamente: {e}")


def corregir_y_resaltar_ortografia(p, idx, alertas_ortografia):
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
            for run in p.runs:
                if re.search(patron, run.text, re.IGNORECASE):
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            texto_original = re.sub(patron, reemplazo, texto_original, flags=re.IGNORECASE)

    if hubo_cambio:
        p.text = texto_original
        for run in p.runs:
            for _, reemplazo in correcciones.items():
                if reemplazo.lower() in run.text.lower():
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        alertas_ortografia.append(f"Parrafo {idx}: Se subrayaron y corrigieron faltas de acentuacion tecnica.")


st.set_page_config(page_title="Auditor Pericial Integral", page_icon="⚖️", layout="centered")

st.title("Auditor Pericial de Formalidad y Sintaxis")
st.write("Sube el PDF de Solicitud y el Word del Dictamen para ejecutar la auditoria.")

st.subheader("1. Carga de Documentos Oficiales")
col_pdf, col_docx = st.columns(2)

with col_pdf:
    archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
    archivo_docx = st.file_uploader("Subir Dictamen Pericial en Word (.docx)", type=["docx"])

if archivo_pdf is not None and archivo_docx is not None:
    st.info("Archivos recibidos. Procesando auditoria...")

    texto_pdf = ""
    try:
        archivo_pdf.seek(0)
        bytes_pdf = archivo_pdf.read()

        with pdfplumber.open(io.BytesIO(bytes_pdf)) as pdf:
            for pagina in pdf.pages:
                t = pagina.extract_text(layout=True)
                if t:
                    texto_pdf += "\n" + t

        if not texto_pdf.strip() and HAS_OCR:
            try:
                imagenes = convert_from_bytes(bytes_pdf)
                for img in imagenes:
                    texto_pdf += "\n" + pytesseract.image_to_string(img, lang="spa")
            except Exception as ocr_err:
                st.warning(f"No se pudo aplicar OCR en el PDF: {ocr_err}")

    except Exception as e:
        st.error(f"Error al leer el archivo PDF: {e}")

    match_carpeta = re.search(r"(carpeta|expediente|causa|cui|ap|ci)\s*[\w\d\.\-/:]+", texto_pdf, re.IGNORECASE)
    carpeta_solicitud = match_carpeta.group(0).upper().strip() if match_carpeta else "NO DETECTADO"

    match_folio = re.search(r"(folio)\s*[\w\d\.\-:]+", texto_pdf, re.IGNORECASE)
    if not match_folio:
        match_folio = re.search(r"\b\d{5,8}\b", texto_pdf)
    folio_solicitud = match_folio.group(0).upper().strip() if match_folio else "NO DETECTADO"

    match_oficio = re.search(r"(oficio|fgr|aic|pfm|uinp|sub)\s*[\w\d\.\-/:]+", texto_pdf, re.IGNORECASE)
    oficio_solicitud = match_oficio.group(0).upper().strip() if match_oficio else "NO DETECTADO"

    match_remitente = re.search(r"(lic\.|mtro\.|mtra\.|dr\.|dra\.|licenciado|licenciada|c\.)\s+([a-záéíóúñ\s]+)", texto_pdf, re.IGNORECASE)
    remitente_solicitud = match_remitente.group(0).strip().upper() if match_remitente else "NO DETECTADO"

    try:
        archivo_docx.seek(0)
        doc = docx.Document(archivo_docx)
        activar_control_de_cambios(doc)

        alertas_alineacion = []
        alertas_ortografia = []
        observaciones_cotejo = []

        parrafos = doc.paragraphs
        i = 0
        while i < len(parrafos):
            txt_p = parrafos[i].text.strip()
            txt_limpio_p = quitar_acentos(txt_p.lower()).replace(" ", "")
            if txt_limpio_p == "presente":
                j = i - 1
                while j >= 0 and not parrafos[j].text.strip():
                    p_element = parrafos[j]._element
                    p_element.getparent().remove(p_element)
                    alertas_alineacion.append("Se eliminaron espacios vacios previos a la palabra PRESENTE.")
                    j -= 1
            i += 1

        for idx, p in enumerate(doc.paragraphs, start=1):
            txt = p.text.strip()
            tiene_imagen = len(p._element.xpath('.//w:drawing | .//w:pict')) > 0
            txt_lower = txt.lower()
            txt_limpio = quitar_acentos(txt_lower)

            if txt:
                corregir_y_resaltar_ortografia(p, idx, alertas_ortografia)

            if txt_limpio.replace(" ", "") == "presente":
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_after = Pt(0)
                continue

            if not txt and not tiene_imagen:
                continue

            es_nombre_fotografia = any(txt_limpio.startswith(prefix) for prefix in ["fotografia", "foto", "figura", "imagen", "grafica", "iluminacion"])
            if tiene_imagen or es_nombre_fotografia:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                continue

            es_asunto = "asunto:" in txt_limpio or "se emite dictamen" in txt_limpio
            es_leyenda_oficial = "margarita maza parada" in txt_limpio or "ano de" in txt_limpio
            es_fecha = bool(re.search(r'\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b', txt_limpio)) and len(txt) < 60
            es_derecha = es_asunto or es_leyenda_oficial or es_fecha

            es_palabra_dictamen = "dictamen" in txt_limpio.replace(" ", "") and len(txt) < 30
            es_atentamente = "atentamente" in txt_limpio and len(txt) < 30
            es_perito = "perito en criminalistica" in txt_limpio
            es_centrado = es_palabra_dictamen or es_atentamente or es_perito

            if es_derecha:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif es_centrado:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                if len(txt) > 40:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        st.success("Auditoria completada exitosamente")
        st.divider()

        st.subheader("1. Datos Extraidos del PDF")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Carpeta Inv.", carpeta_solicitud)
        col2.metric("Folio", folio_solicitud)
        col3.metric("Oficio", oficio_solicitud)
        col4.metric("Remitente", remitente_solicitud)

        st.divider()

        st.subheader("2. Errores Gramaticales y Ortograficos Detectados")
        if alertas_ortografia:
            for ao in list(set(alertas_ortografia))[:10]:
                st.warning(ao)
        else:
            st.success("No se detectaron faltas de acentuacion tecnica en palabras clave.")

        st.divider()

        st.subheader("3. Descargar Word Auditado")
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)

        st.download_button(
            label="Descargar Word con Control de Cambios y Resaltado Amarillo",
            data=bio,
            file_name="DICTAMEN_AUDITADO_CONTROL_CAMBIOS.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as doc_err:
        st.error(f"Error procesando el documento Word: {doc_err}")

else:
    st.warning("Por favor, sube ambos archivos para iniciar la auditoria.")
