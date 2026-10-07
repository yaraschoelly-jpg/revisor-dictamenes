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

    hubo_cambio = False

    for run in p.runs:
        texto_run_original = run.text
        if not texto_run_original.strip():
            continue

        for patron, reemplazo in correcciones.items():
            if re.search(patron, texto_run_original, re.IGNORECASE):
                hubo_cambio = True
                nuevo_texto_run = re.sub(patron, reemplazo, texto_run_original, flags=re.IGNORECASE)
                run.text = nuevo_texto_run
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                texto_run_original = nuevo_texto_run

    if hubo_cambio:
        alertas_ortografia.append(f"Parrafo {idx}: Se subrayaron y corrigieron faltas de acentuacion tecnica (conservando negritas).")


def extraer_texto_pdf(archivo_pdf_obj):
    texto_pdf = ""
    try:
        archivo_pdf_obj.seek(0)
        bytes_pdf = archivo_pdf_obj.read()

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
            except Exception:
                pass
    except Exception:
        pass
    return texto_pdf


st.set_page_config(page_title="Auditor Pericial Integral", page_icon="⚖️", layout="centered")

st.title("Auditor Pericial de Formalidad y Sintaxis")
st.write("Sube el PDF de Solicitud, el Word del Dictamen y el PDF doc sop para ejecutar la auditoria.")

st.subheader("1. Carga de Documentos Oficiales")
col_pdf, col_docx, col_sop = st.columns(3)

with col_pdf:
    archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
    archivo_docx = st.file_uploader("Subir Dictamen Pericial en Word (.docx)", type=["docx"])
with col_sop:
    archivo_sop = st.file_uploader("Subir doc sop (PDF)", type=["pdf"])

if archivo_pdf is not None and archivo_docx is not None and archivo_sop is not None:
    st.info("Archivos recibidos. Procesando auditoria...")

    # --- 1. LECTURA DEL PDF DE SOLICITUD ---
    texto_pdf = extraer_texto_pdf(archivo_pdf)

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

    # --- 2. LECTURA Y VALIDACIÓN DEL PDF "DOC SOP" ---
    texto_sop = extraer_texto_pdf(archivo_sop)
    texto_sop_clean = quitar_acentos(texto_sop.lower())

    titulos_sop_requeridos = [
        "comunicacion con la autoridad",
        "documentacion escrita",
        "hoja de datos crudos",
        "identificacion de riesgos",
        "material",
        "procesamiento de indicios",
        "lineas base",
        "croquis"
    ]

    titulos_sop_encontrados = {}
    for tit in titulos_sop_requeridos:
        titulos_sop_encontrados[tit] = tit in texto_sop_clean

    # --- 3. AUDITORÍA EN WORD ---
    try:
        archivo_docx.seek(0)
        doc = docx.Document(archivo_docx)
        activar_control_de_cambios(doc)

        alertas_alineacion = []
        alertas_ortografia = []

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
            es_perito = "perito en criminalistica" in
