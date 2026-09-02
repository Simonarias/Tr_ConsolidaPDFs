import streamlit as st
from pypdf import PdfReader, PdfWriter
import io
import os
import zipfile
import converters

st.set_page_config(
    page_title="Consolidador y Convertidor a PDF",
    layout="centered"
)

st.title("Consolidador y Convertidor a PDF")

tab_convert, tab_merge = st.tabs(["Convertir a PDF", "Consolidar PDFs"])

SUPPORTED_TYPES = [
    "docx", "xlsx", "pptx", 
    "jpg", "jpeg", "png", "webp", "bmp", "tiff", "gif",
    "txt", "csv", "md", "json", "log", "py", "js", "html", "css", "xml"
]

# ==========================================
# PESTAÑA 1: CONVERTIDOR DE ARCHIVOS A PDF
# ==========================================
with tab_convert:
    files_to_convert = st.file_uploader(
        "Selecciona los archivos a convertir",
        type=SUPPORTED_TYPES,
        accept_multiple_files=True,
        key="uploader_convert"
    )
    
    if files_to_convert:
        converted_results = []
        conversion_errors = []
        
        for f in files_to_convert:
            try:
                pdf_bytes = converters.convert_file_to_pdf(f, filename=f.name)
                base_name = os.path.splitext(f.name)[0]
                converted_results.append({
                    "original_name": f.name,
                    "pdf_name": f"{base_name}.pdf",
                    "pdf_bytes": pdf_bytes
                })
            except Exception as e:
                conversion_errors.append((f.name, str(e)))
                
        for err_file, err_msg in conversion_errors:
            st.error(f"Error en '{err_file}': {err_msg}")
            
        if converted_results:
            col_zip, col_merge_all = st.columns(2)
            
            # 1. Descargar todos en ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in converted_results:
                    zf.writestr(item["pdf_name"], item["pdf_bytes"])
            zip_buffer.seek(0)
            
            with col_zip:
                st.download_button(
                    label="Descargar todos (.ZIP)",
                    data=zip_buffer,
                    file_name="archivos_convertidos.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
            # 2. Descargar todos unidos en 1 solo PDF
            writer_all = PdfWriter()
            for item in converted_results:
                reader = PdfReader(io.BytesIO(item["pdf_bytes"]))
                for page in reader.pages:
                    writer_all.add_page(page)
                
            merged_all_buffer = io.BytesIO()
            writer_all.write(merged_all_buffer)
            merged_all_buffer.seek(0)
            
            with col_merge_all:
                st.download_button(
                    label="Unir todos en 1 PDF",
                    data=merged_all_buffer,
                    file_name="todos_convertidos_unidos.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
            st.divider()
            
            for idx, item in enumerate(converted_results):
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.text(item['original_name'])
                with c2:
                    st.download_button(
                        label="Descargar PDF",
                        data=item["pdf_bytes"],
                        file_name=item["pdf_name"],
                        mime="application/pdf",
                        key=f"dl_single_{idx}",
                        use_container_width=True
                    )

# ==========================================
# PESTAÑA 2: CONSOLIDADOR DE PDFS
# ==========================================
with tab_merge:
    MERGE_TYPES = ["pdf"] + SUPPORTED_TYPES
    
    uploaded_merge_files = st.file_uploader(
        "Selecciona los archivos a consolidar",
        type=MERGE_TYPES,
        accept_multiple_files=True,
        key="uploader_merge"
    )
    
    if uploaded_merge_files:
        file_map = {f.name: f for f in uploaded_merge_files}
        file_names = list(file_map.keys())
        
        orden = st.multiselect(
            "Orden final de los documentos",
            options=file_names,
            default=file_names
        )
        
        if orden:
            if st.button("Consolidar en un solo PDF", type="primary", use_container_width=True):
                with st.spinner("Consolidando..."):
                    try:
                        writer = PdfWriter()
                        total_paginas = 0
                        
                        for nombre in orden:
                            archivo = file_map[nombre]
                            pdf_bytes = converters.convert_file_to_pdf(archivo, filename=nombre)
                            pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
                            for page in pdf_reader.pages:
                                writer.add_page(page)
                            total_paginas += len(pdf_reader.pages)
                        
                        output_pdf = io.BytesIO()
                        writer.write(output_pdf)
                        output_pdf.seek(0)
                        
                        if total_paginas > 0:
                            st.download_button(
                                label="Descargar PDF Consolidado",
                                data=output_pdf,
                                file_name="consolidado.pdf",
                                mime="application/pdf",
                                type="primary",
                                use_container_width=True
                            )
                    except Exception as e:
                        st.error(f"Error al consolidar: {e}")
