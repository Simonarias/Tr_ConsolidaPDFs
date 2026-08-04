import streamlit as st
from pypdf import PdfReader, PdfWriter
import io
import os
import zipfile
import converters

# Configuración de la página
st.set_page_config(
    page_title="Consolidador & Convertidor a PDF",
    page_icon="📄",
    layout="centered"
)

# Estilos personalizados premium con gradientes y tarjetas interactivas
st.markdown("""
    <style>
    .main {
        padding: 1.5rem;
    }
    .stAlert {
        border-radius: 10px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        background-image: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%);
        color: white;
        font-weight: 600;
        border: none;
        padding: 0.6rem 1rem;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(124, 58, 237, 0.4);
        color: white;
    }
    .main-header {
        text-align: center;
        margin-bottom: 2rem;
    }
    .main-title {
        background: -webkit-linear-gradient(#4F46E5, #7C3AED);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        font-size: 2.3rem;
        margin-bottom: 0.5rem;
    }
    .sub-title {
        color: #475569;
        font-size: 1.05rem;
    }
    .file-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1rem;
        margin-bottom: 0.8rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .badge-success {
        background-color: #DEF7EC;
        color: #03543F;
        padding: 0.25rem 0.6rem;
        border-radius: 9999px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

# Encabezado principal
st.markdown("""
    <div class="main-header">
        <div class="main-title">📄 Suite de Herramientas PDF</div>
        <div class="sub-title">Convierte tus archivos (Word, Excel, PowerPoint, Imágenes, Texto) a PDF y consolídalos en un único documento de forma instantánea.</div>
    </div>
""", unsafe_allow_html=True)

# Pestañas principales de la aplicación
tab_convert, tab_merge, tab_info = st.tabs([
    "🔄 Convertir a PDF", 
    "🔀 Consolidar PDFs", 
    "ℹ️ Formatos Soportados"
])

# ==========================================
# PESTAÑA 1: CONVERTIDOR DE ARCHIVOS A PDF
# ==========================================
with tab_convert:
    st.subheader("🔄 Módulo de Conversión a PDF")
    st.write("Sube uno o varios archivos en distintos formatos para convertirlos automáticamente a PDF.")
    
    SUPPORTED_TYPES = [
        "docx", "xlsx", "pptx", 
        "jpg", "jpeg", "png", "webp", "bmp", "tiff", "gif",
        "txt", "csv", "md", "json", "log", "py", "js", "html", "css", "xml"
    ]
    
    files_to_convert = st.file_uploader(
        "Selecciona los archivos a convertir:",
        type=SUPPORTED_TYPES,
        accept_multiple_files=True,
        key="uploader_convert"
    )
    
    if files_to_convert:
        st.success(f"📂 Has cargado {len(files_to_convert)} archivo(s) para conversión.")
        
        converted_results = []
        conversion_errors = []
        
        with st.spinner("Procesando y convirtiendo archivos a PDF..."):
            for f in files_to_convert:
                try:
                    pdf_io = converters.convert_file_to_pdf(f, filename=f.name)
                    base_name = os.path.splitext(f.name)[0]
                    pdf_filename = f"{base_name}.pdf"
                    converted_results.append({
                        "original_name": f.name,
                        "pdf_name": pdf_filename,
                        "pdf_bytes": pdf_io.getvalue()
                    })
                except Exception as e:
                    conversion_errors.append((f.name, str(e)))
                    
        if conversion_errors:
            for err_file, err_msg in conversion_errors:
                st.error(f"❌ Error al convertir **{err_file}**: {err_msg}")
                
        if converted_results:
            st.markdown("### 📋 Archivos Convertidos")
            
            # Opciones globales de descarga
            col_zip, col_merge_all = st.columns(2)
            
            # 1. Descargar todos en ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in converted_results:
                    zf.writestr(item["pdf_name"], item["pdf_bytes"])
            zip_buffer.seek(0)
            
            with col_zip:
                st.download_button(
                    label="📦 Descargar todos en ZIP",
                    data=zip_buffer,
                    file_name="archivos_convertidos.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
            # 2. Descargar todos consolidados en 1 solo PDF
            writer_all = PdfWriter()
            total_pages = 0
            for item in converted_results:
                reader = PdfReader(io.BytesIO(item["pdf_bytes"]))
                for page in reader.pages:
                    writer_all.add_page(page)
                total_pages += len(reader.pages)
                
            merged_all_buffer = io.BytesIO()
            writer_all.write(merged_all_buffer)
            merged_all_buffer.seek(0)
            
            with col_merge_all:
                st.download_button(
                    label="📄 Unir todos en 1 solo PDF",
                    data=merged_all_buffer,
                    file_name="todos_convertidos_unidos.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
            st.divider()
            
            # Lista individual de descargas
            for idx, item in enumerate(converted_results):
                c1, c2, c3 = st.columns([3, 2, 2])
                with c1:
                    st.write(f"📄 **{item['original_name']}**")
                with c2:
                    st.markdown("<span class='badge-success'>🟢 Convertido a PDF</span>", unsafe_allow_html=True)
                with c3:
                    st.download_button(
                        label="⬇️ Descargar PDF",
                        data=item["pdf_bytes"],
                        file_name=item["pdf_name"],
                        mime="application/pdf",
                        key=f"dl_single_{idx}"
                    )
    else:
        st.info("💡 Sube tus archivos arriba para convertirlos a PDF en segundos.")

# ==========================================
# PESTAÑA 2: CONSOLIDADOR DE PDFS Y OTROS FORMATOS
# ==========================================
with tab_merge:
    st.subheader("🔀 Módulo de Consolidador de PDFs")
    st.write("Sube múltiples archivos PDF o de otros formatos para ordenarlos y combinarlos en un solo archivo PDF final.")
    
    MERGE_TYPES = ["pdf"] + SUPPORTED_TYPES
    
    uploaded_merge_files = st.file_uploader(
        "Selecciona los archivos a fusionar/consolidar:",
        type=MERGE_TYPES,
        accept_multiple_files=True,
        key="uploader_merge"
    )
    
    if uploaded_merge_files:
        st.info(f"📁 Has cargado {len(uploaded_merge_files)} archivo(s).")
        
        # Mapeo de nombres a archivos
        file_map = {f.name: f for f in uploaded_merge_files}
        file_names = list(file_map.keys())
        
        st.subheader("🔄 Orden de consolidación")
        st.write("Organiza el orden en el que quieres que aparezcan los documentos en el PDF final:")
        
        orden = st.multiselect(
            "Arrastra o selecciona los archivos en el orden deseado:",
            options=file_names,
            default=file_names,
            help="El primer elemento de la lista será la primera página, el segundo a continuación, y así sucesivamente."
        )
        
        if not orden:
            st.warning("⚠️ Debes seleccionar al menos un archivo en la lista de ordenamiento.")
        else:
            if st.button("✨ Consolidar en un solo PDF"):
                with st.spinner("Procesando, convirtiendo si es necesario y consolidando archivos..."):
                    try:
                        writer = PdfWriter()
                        total_paginas = 0
                        
                        for nombre in orden:
                            archivo = file_map[nombre]
                            # Convertir a PDF si no lo es
                            pdf_buffer = converters.convert_file_to_pdf(archivo, filename=nombre)
                            pdf_reader = PdfReader(pdf_buffer)
                            for page in pdf_reader.pages:
                                writer.add_page(page)
                            total_paginas += len(pdf_reader.pages)
                        
                        output_pdf = io.BytesIO()
                        writer.write(output_pdf)
                        output_pdf.seek(0)
                        
                        if total_paginas == 0:
                            st.error("❌ Ningún archivo válido pudo ser procesado.")
                        else:
                            st.success(f"🎉 ¡Consolidación exitosa! Total de páginas generadas: {total_paginas}")
                            
                            st.download_button(
                                label="⬇️ Descargar PDF Consolidado",
                                data=output_pdf,
                                file_name="consolidado.pdf",
                                mime="application/pdf"
                            )
                    except Exception as e:
                        st.error(f"❌ Ocurrió un error al fusionar los archivos: {e}")
    else:
        st.info("💡 Sube tus archivos PDF u otros formatos en la parte superior para comenzar.")

# ==========================================
# PESTAÑA 3: FORMATOS SOPORTADOS E INFORMACIÓN
# ==========================================
with tab_info:
    st.subheader("ℹ️ Formatos Soportados para Conversión a PDF")
    
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.markdown("""
        #### 🖼️ Imágenes
        * **Formatos:** `.jpg`, `.jpeg`, `.png`, `.webp`, `.bmp`, `.tiff`, `.gif`
        * **Resultado:** Se convierten a páginas PDF manteniendo su resolución y calidad.
        
        #### 📝 Documentos de Word
        * **Formatos:** `.docx`
        * **Resultado:** Extrae texto, encabezados y tablas formateadas a un documento PDF estructurado.
        
        #### 📊 Hojas de Cálculo (Excel)
        * **Formatos:** `.xlsx`
        * **Resultado:** Transforma cada hoja de cálculo en tablas PDF organizadas en formato apaisado.
        """)
        
    with col_b:
        st.markdown("""
        #### 🖥️ Presentaciones (PowerPoint)
        * **Formatos:** `.pptx`
        * **Resultado:** Convierte diapositivas con texto e información a páginas PDF.
        
        #### 📄 Archivos de Texto y Código
        * **Formatos:** `.txt`, `.csv`, `.md`, `.json`, `.log`, `.py`, `.js`, `.html`, `.css`, `.xml`
        * **Resultado:** Formatea el código o texto en un diseño de PDF monoespaciado y limpio.
        """)
    
    st.divider()
    st.markdown("🔒 **Seguridad y Privacidad:** Todos los archivos se procesan 100% en memoria en la sesión de tu navegador y no se guardan en el servidor.")
