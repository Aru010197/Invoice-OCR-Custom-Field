import streamlit as st
from PIL import Image
from pdf2image import convert_from_path, convert_from_bytes
import pytesseract
import os
import io
import pandas as pd
import xlsxwriter
import numpy as np
import cv2
import re
from pathlib import Path
import traceback
from streamlit_drawable_canvas import st_canvas

# --- Helper Functions (adapted from InvoiceOCR class and general utility) ---

def log_message(message, level="info"):
    """Helper to display messages in Streamlit."""
    if level == "info":
        st.info(message)
    elif level == "success":
        st.success(message)
    elif level == "warning":
        st.warning(message)
    elif level == "error":
        st.error(message)
    else:
        st.write(message)

def pdf_to_images_st(pdf_bytes, dpi=300):
    """Convert PDF bytes to a list of PIL Images."""
    try:
        images = convert_from_bytes(pdf_bytes, dpi=dpi, poppler_path=None)
        log_message(f"Converted {len(images)} pages from PDF.", "info")
        return images
    except Exception as e:
        log_message(f"Error converting PDF to images: {e}", "error")
        st.error(f"Detailed PDF conversion error: {traceback.format_exc()}")
        # Attempt to provide more specific advice if poppler is the issue
        if "Poppler" in str(e) or "pdfinfo" in str(e):
            st.error("This error might be related to Poppler. Ensure Poppler is installed and in your system's PATH.")
            st.error("For Windows, download Poppler binaries and add the 'bin' folder to PATH.")
            st.error("For macOS (using Homebrew): brew install poppler")
            st.error("For Linux (Debian/Ubuntu): sudo apt-get install poppler-utils")
        return []

def preprocess_image_st(pil_image):
    """Preprocess PIL image for OCR (e.g., grayscale).
       Note: More advanced preprocessing like adaptiveThreshold might be added later.
    """
    log_message("Preprocessing image...", "info")
    try:
        image_cv = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
        # For now, just return grayscale PIL Image
        processed_pil_image = Image.fromarray(gray)
        log_message("Image preprocessing (to grayscale) complete.", "info")
        return processed_pil_image
    except Exception as e:
        log_message(f"Error during image preprocessing: {e}", "error")
        return pil_image # Return original if error

def extract_text_from_image_st(pil_image, tesseract_config='--oem 3 --psm 6'):
    """Extract text from a PIL image using Tesseract."""
    log_message("Extracting text from image...", "info")
    try:
        text = pytesseract.image_to_string(pil_image, config=tesseract_config)
        log_message("Text extraction successful.", "info")
        return text.strip()
    except Exception as e:
        log_message(f"Error during Tesseract OCR: {e}", "error")
        if "tesseract is not installed" in str(e).lower() or "tesseractnotfound" in str(e).lower():
            st.error("Tesseract OCR is not installed or not found in your system's PATH.")
            st.error("Please install Tesseract OCR and add it to your PATH.")
            st.error("See: https://tesseract-ocr.github.io/tessdoc/Installation.html")
        return ""

def export_fields_to_excel_st(field_ocr_results):
    """Export OCRed field data to an Excel file in memory (bytes).
       Field names are column headers. Multi-line values for a field are listed in subsequent rows under that field's column.
    """
    log_message("Exporting field data to Excel (vertical multi-line format)...", "info")
    output = io.BytesIO()
    try:
        with xlsxwriter.Workbook(output, {'in_memory': True}) as workbook:
            worksheet = workbook.add_worksheet('Extracted Fields')
            
            if not field_ocr_results:
                worksheet.write(0, 0, "No field regions were defined or processed.")
                log_message("No field data to export.", "info")
                return output.getvalue()

            field_names = list(field_ocr_results.keys())
            
            # Prepare data: split multi-line strings into lists of strings
            # and find the maximum number of lines any field has.
            parsed_field_data = []
            max_lines = 0
            for field_name in field_names:
                extracted_text = field_ocr_results.get(field_name, "")
                lines = []
                if extracted_text and isinstance(extracted_text, str):
                    lines = [line for line in extracted_text.split('\n') if line.strip()] # Use '\n' for split
                    if not lines and extracted_text.strip(): # Text was e.g. "   " or "  value  "
                        lines = [extracted_text.strip()]
                elif extracted_text is None:
                    lines = [""] # Represent None as a single empty string entry
                else: # Not a string or empty string
                    lines = [str(extracted_text)]
                
                if not lines: # If after all processing, lines is empty (e.g. original was empty string)
                    lines = [""] # Ensure at least one (empty) entry to maintain structure

                parsed_field_data.append(lines)
                if len(lines) > max_lines:
                    max_lines = len(lines)
            
            # Write headers (field names) in the first row (row 0)
            for col_idx, name in enumerate(field_names):
                worksheet.write(0, col_idx, name)
            
            # Write data rows
            # max_lines will be 0 if field_ocr_results was empty or all values were empty strings that got filtered
            # but we ensured parsed_field_data has at least one [""] for each field if it was originally empty.
            # So max_lines will be at least 1 if there were any fields.
            if max_lines == 0 and field_names: # All fields were empty strings
                max_lines = 1

            for row_num_in_data in range(max_lines):
                excel_row_idx = row_num_in_data + 1 # Start data from Excel row 1
                for col_idx, lines_for_this_field in enumerate(parsed_field_data):
                    if row_num_in_data < len(lines_for_this_field):
                        worksheet.write(excel_row_idx, col_idx, lines_for_this_field[row_num_in_data])
                    else:
                        worksheet.write(excel_row_idx, col_idx, "") # Fill with empty string if this field has fewer lines
            
            # Adjust column widths
            for col_idx, name in enumerate(field_names):
                # Consider header length
                max_col_width = len(name)
                # Consider data lengths in this column
                for lines_for_this_field in parsed_field_data:
                    if col_idx < len(parsed_field_data): # Should always be true if logic is correct
                        for line_val_idx in range(len(parsed_field_data[col_idx])):
                             line_val = parsed_field_data[col_idx][line_val_idx]
                             max_col_width = max(max_col_width, len(str(line_val)))
                
                final_col_width = max(max_col_width, 10) + 2 # Min width 10, add padding
                worksheet.set_column(col_idx, col_idx, final_col_width)

        log_message("Field data Excel export complete (vertical multi-line format).", "success")
        return output.getvalue()
    except Exception as e:
        log_message(f"Error creating field data Excel workbook (vertical multi-line): {e}\n{traceback.format_exc()}", "error")
        return None

# --- Main Streamlit Application ---
def main():
    st.set_page_config(page_title="Invoice OCR Tool", layout="wide")
    st.title("📄 Invoice OCR Tool")
    st.markdown("Extract data from scanned invoices (PDF, JPG, PNG) and export to Excel.")

    # --- Session State Initialization ---
    if 'selected_fields_template' not in st.session_state:
        st.session_state.selected_fields_template = {}
    if 'processing_done' not in st.session_state:
        st.session_state.processing_done = False
    if 'extracted_data' not in st.session_state:
        st.session_state.extracted_data = {}
    if 'full_text' not in st.session_state:
        st.session_state.full_text = ""
    if 'pil_images_for_processing' not in st.session_state:
        st.session_state.pil_images_for_processing = []
    if 'current_page_for_canvas' not in st.session_state:
        st.session_state.current_page_for_canvas = 0
    if 'canvas_image_to_draw_on' not in st.session_state:
        st.session_state.canvas_image_to_draw_on = None
    if 'canvas_key' not in st.session_state:
        st.session_state.canvas_key = "canvas_initial"
    if 'main_preview_done' not in st.session_state:
        st.session_state.main_preview_done = False # For the preview in the new sidebar

    # --- Main Area for Inputs and Settings (Previously Sidebar) ---
    st.header("⚙️ Settings & Upload")

    uploaded_file = st.file_uploader("Upload Invoice File", type=["pdf", "jpg", "jpeg", "png"])

    dpi = st.number_input("Processing DPI for PDF conversion", min_value=72, max_value=600, value=300, step=50,
                          help="Higher DPI can improve OCR accuracy but increases processing time.")
    tesseract_config = st.text_input("Tesseract Config", value="--oem 3 --psm 6",
                                     help="Advanced Tesseract OCR configuration. Default is usually fine.")

    st.markdown("---")
    st.subheader("🎨 Define Field Regions Interactively")

    if uploaded_file:
        if not st.session_state.pil_images_for_processing: # Load images if not already loaded
            file_bytes_for_canvas = uploaded_file.getvalue()
            file_ext_for_canvas = Path(uploaded_file.name).suffix.lower()
            try:
                if file_ext_for_canvas == ".pdf":
                    st.session_state.pil_images_for_processing = pdf_to_images_st(file_bytes_for_canvas, dpi=dpi)
                elif file_ext_for_canvas in [".jpg", ".jpeg", ".png"]:
                    img_canvas = Image.open(io.BytesIO(file_bytes_for_canvas))
                    st.session_state.pil_images_for_processing = [img_canvas]
                
                if st.session_state.pil_images_for_processing:
                    if uploaded_file.name != st.session_state.get('last_uploaded_filename_canvas', ''):
                        st.session_state.canvas_key = f"canvas_{uploaded_file.name}_{st.session_state.current_page_for_canvas}"
                        st.session_state.last_uploaded_filename_canvas = uploaded_file.name
                        st.session_state.selected_fields_template = {} 
                        st.session_state.current_page_for_canvas = 0
                        st.session_state.main_preview_done = False # Reset preview flag for new sidebar

            except Exception as e_load_canvas:
                st.error(f"Error preparing image for canvas: {e_load_canvas}")
                st.session_state.pil_images_for_processing = []

        if st.session_state.pil_images_for_processing:
            # Create two columns: col1 for canvas, col2 for field definitions and saving
            col1, col2 = st.columns([2, 1])  # Adjust ratio if needed, e.g., [3,2]

            with col1:
                num_pages = len(st.session_state.pil_images_for_processing)
                if num_pages > 1:
                    prev_page_for_canvas = st.session_state.current_page_for_canvas
                    st.session_state.current_page_for_canvas = st.selectbox(
                        "Select Page to Define Fields", 
                        options=range(num_pages), 
                        index=st.session_state.current_page_for_canvas,
                        format_func=lambda x: f"Page {x+1}",
                        key="page_selector_main_area_col1" # Ensure key is unique if necessary
                    )
                    if prev_page_for_canvas != st.session_state.current_page_for_canvas:
                        st.session_state.canvas_key = f"canvas_{uploaded_file.name}_{st.session_state.current_page_for_canvas}"

                current_pil_image = st.session_state.pil_images_for_processing[st.session_state.current_page_for_canvas]
                st.session_state.canvas_image_to_draw_on = current_pil_image.copy()

                # --- DEBUGGING STEP: Display the image using st.image ---
                if st.session_state.canvas_image_to_draw_on:
                    st.write("Debug: Image for canvas background")
                    # Add a unique key to st.image to force re-render
                    st.image(st.session_state.canvas_image_to_draw_on, caption="Image to be used on canvas", use_column_width=True, key=f"debug_img_page_{st.session_state.current_page_for_canvas}")
                    log_message(f"Debug: Canvas image mode: {st.session_state.canvas_image_to_draw_on.mode}, size: {st.session_state.canvas_image_to_draw_on.size}", "info")
                # --- END DEBUGGING STEP ---

                st.write(f"Draw rectangles on Page {st.session_state.current_page_for_canvas + 1} below:")
                
                img_width, img_height = st.session_state.canvas_image_to_draw_on.size
                canvas_width_main = 800 
                canvas_height_main = int(canvas_width_main * (img_height / img_width))
                
                canvas_result = st_canvas(
                    fill_color="rgba(255, 165, 0, 0.0)",
                    stroke_width=2,
                    stroke_color="rgba(255, 0, 0, 1)",
                    background_image=st.session_state.canvas_image_to_draw_on,
                    update_streamlit=True,
                    height=canvas_height_main,
                    width=canvas_width_main,
                    drawing_mode="rect",
                    initial_drawing={'version': '5.3.0', 'objects': []},
                    key=st.session_state.canvas_key
                )

            with col2:
                if canvas_result.json_data is not None and canvas_result.json_data["objects"]:
                    st.markdown("**Defined Fields on Current Page:**")
                    page_fields = {k:v for k,v in st.session_state.selected_fields_template.items() if v['page_index'] == st.session_state.current_page_for_canvas}
                    if page_fields:
                        for name, data in page_fields.items():
                            st.text(f"- {name}: BBox {data['bbox']}")
                    else:
                        st.caption("No fields defined for this page yet.")

                    st.markdown("---")
                    st.write("After drawing a rectangle, name and save it:")
                    field_name_input = st.text_input("Enter Field Name for Last Drawn Rectangle", key=f"field_name_{st.session_state.canvas_key}")

                    if st.button("Save Last Drawn Field", key=f"save_field_{st.session_state.canvas_key}"):
                        log_message(f"Save button clicked. Field name input: '{field_name_input}'", "info") # Changed to info for visibility
                        if canvas_result.json_data and canvas_result.json_data.get("objects"):
                            log_message(f"Canvas JSON data objects: {canvas_result.json_data['objects']}", "info") # Changed to info
                        else:
                            log_message("Canvas JSON data is None or has no objects upon save click.", "warning")

                        if field_name_input and canvas_result.json_data and canvas_result.json_data.get("objects"):
                            last_object = canvas_result.json_data["objects"][-1]
                            log_message(f"Last object from canvas: {last_object}", "info") # Changed to info
                            if last_object["type"] == "rect":
                                orig_w, orig_h = st.session_state.canvas_image_to_draw_on.size
                                # Use canvas_width_main, canvas_height_main from col1 for scaling
                                display_w, display_h = canvas_width_main, canvas_height_main 
                                scale_x = orig_w / display_w
                                scale_y = orig_h / display_h
                                x1 = int(last_object["left"] * scale_x)
                                y1 = int(last_object["top"] * scale_y)
                                x2 = int((last_object["left"] + last_object["width"]) * scale_x)
                                y2 = int((last_object["top"] + last_object["height"]) * scale_y)
                                final_x1, final_x2 = min(x1, x2), max(x1, x2)
                                final_y1, final_y2 = min(y1, y2), max(y1, y2)
                                bbox = [final_x1, final_y1, final_x2, final_y2]
                                st.session_state.selected_fields_template[field_name_input] = {
                                    "page_index": st.session_state.current_page_for_canvas,
                                    "bbox": bbox
                                }
                                log_message(f"Field '{field_name_input}' saved for page {st.session_state.current_page_for_canvas + 1} with BBox: {bbox}", "success")
                                st.experimental_rerun()
                        elif not field_name_input:
                            st.warning("Please enter a name for the field.")
                        elif not (canvas_result.json_data and canvas_result.json_data.get("objects")):
                            st.warning("Please draw a rectangle on the canvas first. No objects found in canvas data.")
                        else:
                            st.warning("Please draw a rectangle and enter a field name.")
                
                if st.session_state.selected_fields_template:
                    st.markdown("---")
                    st.markdown("**All Defined Fields (Across Pages):**")
                    for name, data in st.session_state.selected_fields_template.items():
                        st.text(f"- {name}: Page {data['page_index'] + 1}, BBox {data['bbox']}")
            
            # --- Process Button (Moved to Main Area) ---
            # This button remains outside the columns but within the 'if st.session_state.pil_images_for_processing:' block
            st.markdown("---") 

            # Create columns to position the button on the right side of the main content area
            # The first column acts as a spacer, the second will contain the button.
            # Adjust the ratio [4,1] as needed (e.g., [3,1] for less space on left, [5,1] for more).
            _ , button_col = st.columns([4, 1]) 

            with button_col:
                if st.button("🚀 Process Invoice"): 
                    if not st.session_state.pil_images_for_processing:
                        file_bytes_process = uploaded_file.getvalue()
                        file_ext_process = Path(uploaded_file.name).suffix.lower()
                        if file_ext_process == ".pdf":
                            st.session_state.pil_images_for_processing = pdf_to_images_st(file_bytes_process, dpi=dpi)
                        elif file_ext_process in [".jpg", ".jpeg", ".png"]:
                            img_process = Image.open(io.BytesIO(file_bytes_process))
                            st.session_state.pil_images_for_processing = [img_process]

                    if not st.session_state.pil_images_for_processing:
                        log_message("No image available for processing. Upload a valid file.", "error")
                    else:
                        with st.spinner("Processing invoice... This may take a moment."):
                            log_message("Starting invoice processing...", "info")
                            field_ocr_results = {}
                            full_extracted_text_parts = []

                            if st.session_state.selected_fields_template:
                                log_message("Processing defined field regions...", "info")
                                for page_idx, pil_img in enumerate(st.session_state.pil_images_for_processing):
                                    for field_name, region_data in st.session_state.selected_fields_template.items():
                                        if region_data.get("page_index") == page_idx:
                                            bbox = region_data["bbox"]
                                            try:
                                                cropped_pil_img = pil_img.crop(bbox)
                                                text = extract_text_from_image_st(cropped_pil_img, tesseract_config)
                                                field_ocr_results[field_name] = text
                                                log_message(f"Extracted for '{field_name}': '{text[:50]}...'", "info")
                                            except Exception as e:
                                                log_message(f"Error extracting field '{field_name}': {e}", "error")
                                                field_ocr_results[field_name] = f"ERROR: {e}"
                            st.session_state.extracted_data = field_ocr_results
                            log_message("Field region extraction complete.", "success")

                            log_message("Extracting full text from document...", "info")
                            for i, pil_img_proc in enumerate(st.session_state.pil_images_for_processing):
                                log_message(f"Performing full OCR on page {i+1}...", "info")
                                page_text = extract_text_from_image_st(pil_img_proc, tesseract_config)
                                full_extracted_text_parts.append(page_text)
                            
                            st.session_state.full_text = "\\n--- Page Break ---\\n".join(full_extracted_text_parts)
                            log_message("Full text extraction complete.", "success")
                            
                            st.session_state.processing_done = True
                            log_message("Processing complete!", "success")
                            st.balloons()
                        st.experimental_rerun()

        else: # if not st.session_state.pil_images_for_processing (after upload check)
            st.info("Image/PDF loaded, but no pages found or error during load. Please check file.")
    else: # if not uploaded_file:
        st.info("Upload an image/PDF to begin.")


    # --- Sidebar for Preview and Results (Previously Main Area) ---
    with st.sidebar:
        st.header("📄 Preview & Results")
        if uploaded_file:
            if not st.session_state.get('main_preview_done', False) and st.session_state.pil_images_for_processing:
                st.subheader("File Preview (First Page)")
                try:
                    st.image(st.session_state.pil_images_for_processing[0], caption=f"First page of {uploaded_file.name}", use_column_width=True)
                    st.session_state.main_preview_done = True
                except Exception as e_sidebar_preview:
                    log_message(f"Error loading sidebar preview: {e_sidebar_preview}", "error")
            
            if st.session_state.processing_done:
                st.subheader("📊 Extracted Data")
                if st.session_state.extracted_data:
                    st.markdown("#### Field-Specific Extractions:")
                    log_message(f"Data for Excel export: {st.session_state.extracted_data}", "info") # Added log
                    df_fields = pd.DataFrame(list(st.session_state.extracted_data.items()), columns=['Field Name', 'Extracted Value'])
                    st.dataframe(df_fields)
                    excel_bytes = export_fields_to_excel_st(st.session_state.extracted_data)
                    if excel_bytes: # Added log for success
                        log_message(f"Excel bytes generated. Length: {len(excel_bytes)} bytes.", "success")
                        st.download_button(
                            label="📥 Download Field Data as Excel",
                            data=excel_bytes,
                            file_name=f"{Path(uploaded_file.name).stem}_fields.xlsx" if uploaded_file else "extracted_fields.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_excel_sidebar"
                        )
                    else: # Added log for failure/None
                        log_message("Excel bytes generation failed or returned None.", "error")
                else:
                    st.info("No specific fields were defined or extracted (for Excel export).")

                if st.session_state.full_text:
                    st.markdown("#### Full Extracted Text:")
                    with st.expander("View Full Text", expanded=False):
                        st.text_area("Full Text", st.session_state.full_text, height=200, key="full_text_sidebar")
                    text_bytes = st.session_state.full_text.encode('utf-8')
                    st.download_button(
                        label="📥 Download Full Text as TXT",
                        data=text_bytes,
                        file_name=f"{Path(uploaded_file.name).stem}_fulltext.txt" if uploaded_file else "full_text.txt",
                        mime="text/plain",
                        key="download_txt_sidebar"
                    )
                else:
                    st.info("No full text was extracted.")
            elif uploaded_file and not st.session_state.pil_images_for_processing:
                 st.warning("Could not load the uploaded file for processing or preview.")
            elif not uploaded_file:
                 st.info("Upload a file to see results here.")
        else:
            st.info("Upload a file using the main panel to see preview and results here.")

        st.markdown("---")
        st.caption("Streamlit OCR App v0.1.1")

if __name__ == "__main__":
    main()
