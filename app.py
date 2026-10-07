import io
import re
import unicodedata
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
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


def crear_nodo_del(texto_borrado, autor="Auditor Pericial"):
    """Crea la etiqueta XML nativa w:del (texto tachado/eliminado en Control de Cambios)."""
    del_elem = OxmlElement("w:del")
    del_elem.set(qn("w:id"), "0")
    del_elem.set(qn("w:author"), autor)
    
    r_elem = OxmlElement("w:r")
    del_text_elem = OxmlElement("w:delText")
    del_text_elem.set(qn("xml:space"), "preserve")
    del_text_elem.text = texto_borrado
    
    r_elem.append(del_text_elem)
    del_elem.append(r_elem)
    return del_elem


def crear_nodo_ins(texto_insertado, resaltado_amarillo=True, autor="Auditor Pericial"):
    """Crea la etiqueta XML nativa w:ins (texto insertado con resaltado amarillo en Control de Cambios)."""
    ins_elem = OxmlElement("w:ins")
    ins_elem.set(qn("w:id"), "0")
    ins_elem.set(qn("w:author"), autor)
    
    r_elem = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    
    if resaltado_amarillo:
        highlight = OxmlElement("w:highlight")
        highlight.set(qn("w:val"), "yellow")
        rPr.append(highlight)
        
    r_elem.append(rPr)
    
    t_elem = OxmlElement("w:t")
    t_elem.set(qn("xml:space"), "preserve")
    t_elem.text = texto_insertado
    
    r_elem.append(t_elem)
    ins_elem.append(r_elem)
    return ins_elem


def aplicar_correccion_xml_parrafo(p, diccionario_errores, autor="Auditor Pericial"):
    """
    Recorre los runs del párrafo. Si encuentra un error ortográfico, reemplaza
    el nodo del run por nodos nativos XML <w:del> (palabra original) y <w:ins> (palabra corregida).
    """
    hubo_cambio = False
    p_element = p._element

    for run in list(p.runs):
        texto_run = run.text
        if not texto_run or not texto_run.strip():
            continue

        # Evitar modificar texto que esté entre comillas
        if (texto_run.startswith('"') and texto_run.endswith('"')) or \
           (texto_run.startswith('“') and texto_run.endswith('”')):
            continue

        for patron, reemplazo in diccionario_errores.items():
            match = re.search(patron, texto_run, re.IGNORECASE)
            if match:
                palabra_original = match.group(0)
                
                # Ajuste de mayúsculas/minúsculas
                if palabra_original.isupper():
                    palabra_corregida = reemplazo.upper()
                elif palabra_original[0].isupper():
                    palabra_corregida = reemplazo.capitalize()
                else:
                    palabra_corregida = reemplazo.lower()

                # Dividir el texto antes y después del error
                inicio_idx = match.start()
                fin_idx = match.end()
                
                texto_antes = texto_run[:inicio_idx]
                texto_despues = texto_run[fin_idx:]

                run_node = run._element
                
                # Insertar texto previo si existe
                if texto_antes:
                    run_pre = p.add_run(texto_antes)._element
                    p_element.insert(p_element.index(run_node), run_pre)

                # Insertar los nodos XML nativos de Track Changes
                nodo_del = crear_nodo_del(palabra_original, autor=autor)
                nodo_ins = crear_nodo_ins(palabra_corregida, resaltado_amarillo=True, autor=autor)
                
                p_element.insert(p_element.index(run_node), nodo_del)
                p_element.insert(p_element.index(run_node), nodo_ins)

                # Insertar texto posterior si existe
                if texto_despues:
                    run_post = p.add_run(texto_despues)._element
                    p_element.insert(p_element.index(run_node), run_post)

                # Remover el nodo original reemplazado
                p_element.remove(run_node)
                hubo_cambio = True
                break

    return hubo_cambio


def audit_y_resaltar_parrafo(p, idx, alertas_ortografia):
    txt_p = p.text.strip()
    if not txt_p:
        return

    # Si todo el párrafo está entre comillas, no evaluar ortografía
    if (txt_p.startswith('"') and txt_p.endswith('"')) or \
       (txt_p.startswith('“') and txt_p.endswith('”')) or \
       (txt_p.startswith("'") and txt_p.endswith("'")):
        return

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

    txt_clean = quitar
