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
    """Activa de forma nativa la función de Track Changes (Control de Cambios) en Word."""
    try:
        settings = doc.settings._element
        track_revisions = settings.find(qn("w:trackRevisions"))
        if track_revisions is None:
            track_revisions = OxmlElement("w:trackRevisions")
            settings.append(track_revisions)
    except Exception as e:
        st.warning(f"No se pudo activar el control de cambios automáticamente: {e}")


def audit_y_resaltar_parrafo(p, idx, alertas_ortografia):
    # 1. Regla de Títulos Principales en MAYÚSCULAS Y NEGRITAS
    titulos_mayusculas = {
        r"^\s*antecedentes\b\.?": "ANTECEDENTES.",
        r"^\s*planteamiento\s+del\s+problema\b\.?": "PLANTEAMIENTO DEL PROBLEMA.",
        r"^\s*tecnica\s+de\s+estudio\b\.?": "TÉCNICA DE ESTUDIO.",
        r"^\s*estudio\s+de\s+campo\b\.?": "ESTUDIO DE CAMPO.",
        r"^\s*observacion\b\.?": "OBSERVACIÓN.",
        r"^\s*observaciones\b\.?": "OBSERVACIONES.",
        r"^\s*consideracion\b\.?": "CONSIDERACIÓN.",
        r"^\s*consideraciones\b\.?": "CONSIDERACIONES.",
        r"^\s*conclusion\b\.?": "CONCLUSIÓN.",
        r"^\s*conclusiones\b\.?": "CONCLUSIONES.",
    }

    # 2. Regla de Subtítulos (Primera mayúscula, resto minúsculas, en NEGRITAS)
    subtitulos_tipo_oracion = {
        r"^\s*direccion\b:?": "Dirección:",
        r"^\s*descripcion\s+del\s+lugar\b:?": "Descripción del lugar:",
        r"^\s*examen\s+del\s+lugar\b:?": "Examen del lugar:",
    }

    txt_p = p.text.strip()
    txt_clean = quitar_acentos(txt_p.lower())

    for patron, titulo_correcto in titulos_mayusculas.items():
        if re.search(patron, txt_clean):
            if txt_p != titulo_correcto or any(not r.bold for r in p.runs if r.text.strip()):
                p.text = titulo_correcto
                for r in p.runs:
                    r.bold = True
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
                alertas_ortografia.append(f"Párrafo {idx}: Se formateó el título '{titulo_correcto}' en MAYÚSCULAS y NEGRITAS.")
                return

    for patron, subtitulo_correcto in subtitulos_tipo_oracion.items():
        if re.search(patron, txt_clean):
            if not txt_p.startswith(subtitulo_correcto) or any(not r.bold for r in p.runs if subtitulo_correcto in r.text):
                resto_texto = re.sub(patron, "", txt_p, flags=re.IGNORECASE).strip()
                p.text = ""
                r1 = p.add_run(subtitulo_correcto)
                r1.bold = True
                r1.font.highlight_color = WD_COLOR_INDEX.YELLOW
                if resto_texto:
                    p.add_run(" " + resto_texto)
                alertas_ortografia.append(f"Párrafo {idx}: Se formateó el subtítulo '{subtitulo_correcto}' (Primera mayúscula, resto minúsculas, NEGRITAS).")
                return

    # 3. Corrección Ortográfica Técnica
    diccionario_errores = {
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
        r"\btecnicas\b": "técnicas",
        r"\btecnicos\b": "técnicos",
        r"\bpericiales\b": "periciales",
        r"\banalisis\b": "análisis",
        r"\bfotografia\b": "fotografía",
        r"\bfotografias\b": "fotografías",
        r"\bvehiculo\b": "vehículo",
        r"\bvehiculos\b": "vehículos",
        r"\bmetodologia\b": "metodología",
        r"\bbalistica\b": "balística",
        r"\bdactiloscopia\b": "dactiloscopia"
    }

    hubo_cambio = False

    for run in p.runs:
        texto_run = run.text
        if not texto_run or not texto_run.strip():
            continue

        for patron, reemplazo in diccionario_errores.items():
            if re.search(patron, texto_run, re.IGNORECASE):
                hubo_cambio = True
                texto_run = re.sub(patron, reemplazo, texto_run, flags=re.IGNORECASE)
                run.text = texto_run
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    if hubo_cambio:
        alertas_ortografia.append(f"Párrafo {idx}: Se identificaron correcciones ortográficas y se registraron en Control de Cambios (amarillo).")


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
st.write("Sube el PDF de Solicitud, el Word del Dictamen y el PDF doc sop para ejecutar la auditoría.")

st.subheader("1. Carga de Documentos Oficiales")
col_pdf, col_docx, col_sop = st.columns(3)

with col_pdf:
    archivo_pdf = st.file_uploader("Subir Oficio de Solicitud (PDF)", type=["pdf"])
with col_docx:
    archivo_docx = st.file_uploader("Subir Dictamen Pericial en Word (.docx)", type=["docx"])
with col_sop:
    archivo_sop = st.file_uploader("Subir doc sop (PDF)", type=["pdf"])

if archivo_pdf is not None and archivo_docx is not None and archivo_sop is not None:
    st.info("Archivos recibidos. Procesando auditoría con Control de Cambios activo...")

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

    # --- 2. LECTURA Y VALIDACIÓN DEL PDF "DOC SOP" CON NOMBRES EXACTOS ---
    texto_sop = extraer_texto_pdf(archivo_sop)
    texto_sop_clean = quitar_acentos(texto_sop.lower())

    formatos_sop_requeridos = [
        ("VERIFICACIÓN MATERIAL/ INSUMOS/ EQUIPO PARA INSPECCIONES", r"verificacion\s+material"),
        ("IDENTIFICACIÓN, EVALUACIÓN Y MANEJO DE RIESGOS", r"manejo\s+de\s+riesgos|identificacion.*riesgos"),
        ("DOCUMENTACIÓN ESCRITA", r"documentacion\s+escrita"),
        ("HOJA DE DATOS CRUDOS", r"hoja\s+de\s+datos\s+crudos|datos\s+crudos"),
        ("comunicacion con la autoridad", r"comunicacion\s+con\s+la\s+autoridad"),
        ("croquis simple", r"croquis\s+simple|croquis"),
        ("procesamiento de indicios", r"procesamiento\s+de\s+indicios"),
        ("matriz de datos lineas base", r"matriz\s+de\s+datos\s+lineas\s+base|lineas\s+base")
    ]

    titulos_sop_encontrados = {}
    for nombre_oficial, patron in formatos_sop_requeridos:
        titulos_sop_encontrados[nombre_oficial] = bool(re.search(patron, texto_sop_clean))

    # --- 3. AUDITORÍA EN WORD CON CONTROL DE CAMBIOS ACTIVO ---
    try:
