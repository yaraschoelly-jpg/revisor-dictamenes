import re
import streamlit as st

st.set_page_config(
    page_title="Revisión de Dictámenes y Prompts",
    page_icon="⚖️",
    layout="wide"
)

st.title("⚖️ Sistema de Revisión de Dictámenes y Prompts")
st.markdown("---")

# Lista de palabras esdrújulas frecuentes en el ámbito técnico-pericial
PALABRAS_ESDRUJULAS = [
    "criminalistica", "fisica", "quimica", "balistica", "dactiloscopia",
    "informatica", "caracteristicas", "metodos", "tecnicas", "analisis",
    "elementos", "morfologicos", "dictamenes", "juridico", "tecnico",
    "cientifico", "periciales", "parrafo", "médico", "forense"
]

CORRECCIONES_ESDRUJULAS = {
    "CRIMINALISTICA": "CRIMINALÍSTICA",
    "criminalistica": "criminalística",
    "FISICA": "FÍSICA",
    "fisica": "física",
    "QUIMICA": "QUÍMICA",
    "quimica": "química",
    "BALISTICA": "BALÍSTICA",
    "balistica": "balística",
    "DACTILOSCOPIA": "DACTILOSCOPÍA",
    "dactiloscopia": "dactiloscopía",
    "INFORMATICA": "INFORMÁTICA",
    "informatica": "informática",
    "CARACTERISTICAS": "CARACTERÍSTICAS",
    "caracteristicas": "características",
    "METODOS": "MÉTODOS",
    "metodos": "métodos",
    "TECNICAS": "TÉCNICAS",
    "tecnicas": "técnicas",
    "ANALISIS": "ANÁLISIS",
    "analisis": "análisis",
    "DICTAMENES": "DICTÁMENES",
    "dictamenes": "dictámenes",
    "JURIDICO": "JURÍDICO",
    "juridico": "jurídico",
    "TECNICO": "TÉCNICO",
    "tecnico": "técnico",
    "CIENTIFICO": "CIENTÍFICO",
    "cientifico": "científico",
    "PARRAFO": "PÁRRAFO",
    "parrafo": "párrafo",
}

def procesar_texto(texto: str):
    log_cambios = []
    texto_corregido = texto

    # 1. Corrección de espaciado tras signos de puntuación (comas, puntos, dos puntos)
    # Patrón: un signo de puntuación seguido inmediatamente de una letra o número sin espacio
    patron_espacios = r'([.,;:])([a-zA-ZáéíóúÁÉÍÓÚñÑ0-9])'
    
    def repl_espacio(match):
        signo = match.group(1)
        siguiente = match.group(2)
        log_cambios.append(f"Espacio añadido tras signo de puntuación: `{signo}{siguiente}` ➔ `{signo} {siguiente}`")
        return f"{signo} {siguiente}"

    texto_corregido = re.sub(patron_espacios, repl_espacio, texto_corregido)

    # 2. Corrección de acentuación gráfica en esdrújulas técnicas
    for sin_tilde, con_tilde in CORRECCIONES_ESDRUJULAS.items():
        patron_palabra = r'\b' + re.escape(sin_tilde) + r'\b'
        if re.search(patron_palabra, texto_corregido):
            log_cambios.append(f"Acentuación de esdrújula: `{sin_tilde}` ➔ `{con_tilde}`")
            texto_corregido = re.sub(patron_palabra, con_tilde, texto_corregido)

    # 3. Limpieza sintáctica para Prompts (Asteriscos parásitos dentro de comillas)
    patron_asteriscos_comillas = r'([“"\'«])\s*\*+([^*]+)\*+\s*([”"\'»])'
    if re.search(patron_asteriscos_comillas, texto_corregido):
        log_cambios.append("Limpieza sintáctica de prompt: Eliminación de asteriscos innecesarios dentro de comillas.")
        texto_corregido = re.sub(patron_asteriscos_comillas, r'\1\2\3', texto_corregido)

    # 4. Generación visual de Control de Cambios
    # Marcado visual: ~~Eliminado~~ **Añadido**
    marcado_diff = texto
    # Aplicar reemplazos visuales de espacios
    marcado_diff = re.sub(r'([.,;:])([a-zA-ZáéíóúÁÉÍÓÚñÑ0-9])', r'\1~~ ~~** **\2', marcado_diff)
    
    for sin_tilde, con_tilde in CORRECCIONES_ESDRUJULAS.items():
        patron_palabra = r'\b' + re.escape(sin_tilde) + r'\b'
        marcado_diff = re.sub(patron_palabra, f"~~{sin_tilde}~~ **{con_tilde}**", marcado_diff)

    return texto_corregido, marcado_diff, log_cambios

# Interface Principal
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Texto / Prompt Original")
    input_text = st.text_area(
        "Pega aquí el texto del dictamen o las instrucciones del prompt:",
        height=350,
        placeholder="Ejemplo: Quien suscribe, Lcda. Jennifer Alin Ramírez Pérez,persona perita..."
    )

if st.button("🔍 Ejecutar Revisión y Control de Cambios", type="primary"):
    if not input_text.strip():
        st.warning("Por favor, ingresa un texto para revisar.")
    else:
        texto_limpio, texto_diff, cambios = procesar_texto(input_text)

        with col2:
            st.subheader("📌 Control de Cambios")
            st.markdown(texto_diff)

        st.markdown("---")
        st.subheader("📋 Registro de Observaciones y Correcciones")
        if cambios:
            for cambio in cambios:
                st.write(f"• {cambio}")
        else:
            st.success("No se detectaron errores de espaciado, ortografía o sintaxis en el texto proporcionado.")

        st.subheader("✅ Texto Final Corregido")
        st.code(texto_limpio, language="markdown")
