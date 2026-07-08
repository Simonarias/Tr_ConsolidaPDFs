import streamlit as st
from pypdf import PdfReader, PdfWriter
import io

# Configuración de página con título e icono
st.set_page_config(
    page_title="Consolidador de PDFs",
    page_icon="📄",
    layout="centered"
)

# Estilos personalizados para mejorar la estética (Gradiente premium, bordes redondeados y efectos interactivos)
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stAlert {
        border-radius: 10px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        background-image: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: bold;
        border: none;
        padding: 0.6rem 1rem;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(118, 75, 162, 0.4);
        color: white;
    }
    h1 {
        background: -webkit-linear-gradient(#667eea, #764ba2);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    </style>
""", unsafe_allow_html=True)

st.write("# 📄 Consolidador de Archivos PDF")
st.write("Sube múltiples archivos PDF, ordénalos de forma interactiva y combínalos en un único documento de forma instantánea.")

# Componente de subida de archivos
uploaded_files = st.file_uploader(
    "Selecciona o arrastra tus archivos PDF:",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📁 Has cargado {len(uploaded_files)} archivo(s).")
    
    # Mapeo de nombres a archivos
    file_map = {f.name: f for f in uploaded_files}
    file_names = list(file_map.keys())
    
    # Permitir reordenación visualmente usando un multiselect interactivo
    st.subheader("🔄 Orden de consolidación")
    st.write("Organiza el orden en el que quieres que aparezcan los PDFs en el archivo final:")
    
    orden = st.multiselect(
        "Arrastra o selecciona los archivos en el orden deseado:",
        options=file_names,
        default=file_names,
        help="El primer elemento de la lista será la primera página, el segundo a continuación, y así sucesivamente."
    )
    
    # Validar que al menos haya seleccionado archivos para fusionar
    if not orden:
        st.warning("⚠️ Debes seleccionar al menos un archivo en la lista de ordenamiento.")
    else:
        # Botón para consolidar
        if st.button("✨ Consolidar PDFs"):
            with st.spinner("Procesando y consolidando archivos en memoria..."):
                try:
                    writer = PdfWriter()
                    total_paginas = 0
                    
                    # Consolidar en memoria
                    for nombre in orden:
                        archivo = file_map[nombre]
                        # Leer archivo en memoria
                        pdf_reader = PdfReader(archivo)
                        for page in pdf_reader.pages:
                            writer.add_page(page)
                        total_paginas += len(pdf_reader.pages)
                    
                    # Escribir el resultado en un buffer de bytes en memoria
                    output_pdf = io.BytesIO()
                    writer.write(output_pdf)
                    output_pdf.seek(0)
                    
                    if total_paginas == 0:
                        st.error("❌ Ningún archivo válido pudo ser procesado.")
                    else:
                        st.success(f"🎉 ¡Consolidación exitosa! Total de páginas generadas: {total_paginas}")
                        
                        # Botón de descarga
                        st.download_button(
                            label="⬇️ Descargar PDF Consolidado",
                            data=output_pdf,
                            file_name="consolidado.pdf",
                            mime="application/pdf"
                        )
                except Exception as e:
                    st.error(f"Ocurrió un error al fusionar los PDFs: {e}")
else:
    st.info("💡 Sube tus archivos PDF en la parte superior para comenzar.")
