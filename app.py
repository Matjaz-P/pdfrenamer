import streamlit as st
import fitz  # PyMuPDF
import os
import re
import io
import zipfile
import pandas as pd

# --- Core Processing Function (from your original code, slightly adapted) ---

def extract_table_data(pdf_bytes):
    """Extract table data from PDF bytes in memory."""
    try:
        # Open PDF from bytes
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        for page in doc:
            blocks = page.get_text("dict").get("blocks", [])
            potential_cells = []
            
            for block in blocks:
                if block.get("type") == 0:  # Text block
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            text = span.get("text", "").strip()
                            if text:
                                bbox = span.get("bbox", [])
                                if bbox:
                                    x0, y0, x1, y1 = bbox
                                    potential_cells.append({
                                        'text': text,
                                        'x': x0,
                                        'y': y0
                                    })
            
            # Sort by vertical then horizontal position to read lines correctly
            potential_cells.sort(key=lambda c: (round(c['y'] / 5), c['x']))
            
            # Find the row starting with 'KCZ'
            for i, cell in enumerate(potential_cells):
                if cell['text'] == 'KCZ':
                    row_cells = [cell['text']]
                    current_y = cell['y']
                    
                    # Collect all other cells on the same horizontal line (within a tolerance)
                    for j in range(i + 1, len(potential_cells)):
                        next_cell = potential_cells[j]
                        if abs(next_cell['y'] - current_y) < 5: # 5 pixels tolerance
                            row_cells.append(next_cell['text'])
                        elif next_cell['y'] > current_y + 10: # Moved to the next line
                            break
                    
                    # If we found a full row, return the first 9 elements
                    if len(row_cells) >= 9:
                        doc.close()
                        return row_cells[:9]
        
        doc.close()
        return None
        
    except Exception as e:
        st.error(f"Error processing a PDF file: {e}")
        return None

# --- Streamlit User Interface ---

def main():
    st.set_page_config(page_title="PDF Renamer", layout="wide")

    # --- Header ---
    st.title("📄 PDF Table Extractor & Renamer")
    st.markdown("""
    This tool extracts specific table data from your PDF files to automatically rename them.
    The new name will follow the format: `KCZ_DMX_O_03_TL_PR_004_R_14.pdf`.
    """)

    # --- File Uploader ---
    uploaded_files = st.file_uploader(
        "📁 Select PDF files to process",
        type="pdf",
        accept_multiple_files=True
    )

    if uploaded_files:
        st.info(f"**{len(uploaded_files)}** file(s) selected. Click the button below to start processing.")

        # --- Process Button ---
        if st.button("⚙️ Process & Rename Files", type="primary"):
            
            results = []
            renamed_files_zip_buffer = io.BytesIO()
            zip_archive = zipfile.ZipFile(renamed_files_zip_buffer, 'w', zipfile.ZIP_DEFLATED)
            
            success_count = 0
            fail_count = 0
            
            # Placeholders for progress bar and status text
            progress_bar = st.progress(0, text="Starting...")
            status_text = st.empty()
            
            # Keep track of new names to avoid duplicates within the zip
            generated_new_names = set()

            for i, uploaded_file in enumerate(uploaded_files):
                original_filename = uploaded_file.name
                status_text.text(f"Processing: {original_filename}")
                
                # Read file bytes
                pdf_bytes = uploaded_file.getvalue()
                
                # Extract data
                table_data = extract_table_data(pdf_bytes)
                
                if table_data and len(table_data) >= 9:
                    # Create new name and sanitize it
                    base_name = '_'.join(table_data[:9])
                    new_name = re.sub(r'[<>:"/\\|?*]', '_', base_name) + '.pdf'
                    
                    # Handle duplicate new names
                    counter = 1
                    final_new_name = new_name
                    while final_new_name in generated_new_names:
                        final_new_name = f"{base_name}_{counter}.pdf"
                        counter += 1
                    generated_new_names.add(final_new_name)

                    # Add the file to the zip with the new name
                    zip_archive.writestr(final_new_name, pdf_bytes)
                    
                    results.append({
                        "Status": "✅ Success",
                        "Original Name": original_filename,
                        "New Name": final_new_name
                    })
                    success_count += 1
                else:
                    results.append({
                        "Status": "❌ Failed",
                        "Original Name": original_filename,
                        "New Name": "Could not extract required data."
                    })
                    fail_count += 1
                
                # Update progress bar
                progress_bar.progress((i + 1) / len(uploaded_files), text=f"Processed: {original_filename}")

            status_text.success(f"Processing complete! Success: {success_count}, Failed: {fail_count}")
            progress_bar.progress(1.0)
            
            # --- Display Results Table ---
            st.subheader("📊 Processing Results")
            df = pd.DataFrame(results)

            def style_status(s):
                return ['color: green' if v == '✅ Success' else 'color: red' for v in s]
            
            st.dataframe(df.style.apply(style_status, subset=['Status']), use_container_width=True)

            # --- Create Download Button for the ZIP file ---
            if success_count > 0:
                zip_archive.close()
                st.download_button(
                    label="📥 Download Renamed Files (.zip)",
                    data=renamed_files_zip_buffer.getvalue(),
                    file_name="renamed_pdfs.zip",
                    mime="application/zip",
                )
    else:
        st.info("Please upload one or more PDF files to begin.")

if __name__ == "__main__":
    main()