
# Invoice OCR Tool (Streamlit Version)

A Python utility for converting scanned invoices to structured Excel format with OCR text extraction, now using Streamlit for a web-based interface.

## Overview

This tool allows you to extract data from scanned invoices (PDF, JPG, PNG formats) and export it to Excel or TXT files. It uses OCR to extract text content. The Streamlit interface provides a user-friendly way to upload files and view results.

## Features

- Convert PDF, JPG, and PNG invoice scans to Excel sheets and TXT files
- Extract text using Tesseract OCR
- User-friendly web-based GUI interface powered by Streamlit
- Interactive definition of fields by drawing on the image for targeted data extraction
- Download extracted field data as Excel and full text as TXT

## Requirements

- Python 3.7+
- pdf2image (with Poppler)
- pytesseract (check `requirements.txt` for version)
- OpenCV (for image preprocessing, check `requirements.txt`)
- numpy
- streamlit (version 1.12.0 for compatibility with `streamlit-drawable-canvas`)
- streamlit-drawable-canvas (check `requirements.txt`)
- xlsxwriter
- Pillow
- altair<5 (for Streamlit 1.12.0 compatibility)
- pandas

## Installation

1. Clone or download this repository.
2. Install the required Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Install system dependencies:

   - **Tesseract OCR**:
     - macOS: `brew install tesseract`
     - Ubuntu/Debian: `apt-get install tesseract-ocr`
     - Windows: Download installer from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki) and ensure it's in your PATH.

   - **Poppler** (required for pdf2image):
     - macOS: `brew install poppler`
     - Ubuntu/Debian: `apt-get install poppler-utils`
     - Windows: Download from a reliable source (e.g., [Poppler for Windows by ocr.space](https://ocr.space/poppler)) and add the `bin` directory to your PATH.

## Usage

1.  Navigate to the project directory in your terminal.
2.  Ensure all dependencies from `requirements.txt` are installed.
3.  Run the Streamlit application:
    ```bash
    streamlit run streamlit_app.py
    ```

4.  The application will open in your web browser. Follow these steps:

    **In the Main Area:**
    *   **Upload Invoice:** Use the "Upload Invoice File" uploader to select your PDF, JPG, JPEG, or PNG file.
    *   **Configure Settings (Optional):**
        *   Adjust "Processing DPI for PDF conversion" if needed. Higher DPI can improve OCR but takes longer.
        *   Modify "Tesseract Config" for advanced OCR engine settings if you're familiar with them.
    *   **Define Field Regions Interactively (Optional):**
        *   Once a file is uploaded, an image preview (or the first page of a PDF) will appear on the left side of the "Define Field Regions Interactively" section.
        *   If it's a multi-page PDF, a "Select Page to Define Fields" dropdown will appear above the image. Use this to navigate to the desired page.
        *   Draw a rectangle directly on the image over the area you want to extract.
        *   To the right of the image, you'll see sections for managing fields:
            *   "Defined Fields on Current Page": Lists fields already saved for the page you're viewing.
            *   "After drawing a rectangle, name and save it": Enter a descriptive name for the rectangle you just drew in the "Enter Field Name..." box.
            *   Click "Save Last Drawn Field". The field will be added to the list, and the interface will refresh.
            *   "All Defined Fields (Across Pages)": Shows a summary of all fields you've defined for the entire document.
        *   Repeat this process for all fields you want to extract on any page.
    *   **Process Invoice:** Once you've uploaded a file (and optionally defined fields), click the "🚀 Process Invoice" button located below the field definition area.

    **In the Sidebar (Preview & Results):**
    *   **File Preview:** A preview of the first page of your uploaded file is shown.
    *   **Extracted Data (after processing):**
        *   **Field-Specific Extractions:** If you defined fields, a table with the extracted data for those fields will appear. You can download this data using the "📥 Download Field Data as Excel" button.
        *   **Full Extracted Text:** The complete text extracted from the document will be available. You can view it in an expandable section and download it using the "📥 Download Full Text as TXT" button.

5.  The application will process the document, perform OCR, and display the results as described above.

## How It Works (Streamlit Version)

1. **File Upload**: User uploads an invoice file via the Streamlit interface in the main application area.
2. **Document Preparation**: For PDFs, the tool converts pages to images using `pdf2image`.
3. **Image Display & Interactive Annotation**: Shows the selected page of the document on a canvas in the main area. Users can draw rectangles to define regions of interest.
4. **Field Definition**: Users name the drawn rectangles. These definitions (field name, page, bounding box) are stored.
5. **OCR Processing**: Uses Tesseract (`pytesseract`) to extract text from the full document and/or the user-defined regions.
6. **Data Display and Export**: Shows extracted data in the Streamlit sidebar. Field-specific data can be downloaded as Excel, and full text as a TXT file.

## Troubleshooting

- **PDF Conversion Issues**: Ensure Poppler is correctly installed and its `bin` directory is in your system's PATH. The app will try to provide guidance if it detects Poppler-related errors.
- **OCR Quality Problems**:
    - Ensure Tesseract OCR is installed correctly and accessible via PATH.
    - Try adjusting the DPI setting for PDF conversion (higher DPI might improve quality but takes longer).
    - The quality of the original scan significantly impacts OCR accuracy.
- **`ImportError: dlopen(...) incompatible architecture` (e.g., for Pillow, cv2, numpy, charset-normalizer)**: This usually happens on macOS with ARM (M1/M2/M3) chips if an x86_64 version of a library was installed. The fix is typically to force a reinstall for the correct architecture:
  ```bash
  pip uninstall <package_name> -y
  pip install <package_name> --no-cache-dir --force-reinstall
  # or sometimes just `pip install <package_name>` after uninstalling is enough.
  ```
  Always ensure your pip and Python are targeting the arm64 architecture if on an ARM Mac.
- **`AttributeError: module 'streamlit.elements.image' has no attribute 'image_to_url'`**: This occurs if `streamlit-drawable-canvas` (e.g., v0.9.3) is used with a Streamlit version >= 1.13.0. The solution applied here is to downgrade Streamlit to `1.12.0` as specified in `requirements.txt`.
- **`TypeError: ButtonMixin.button() got an unexpected keyword argument 'type'` or `AttributeError: module 'streamlit' has no attribute 'rerun'` or `TypeError: DataFrameSelectorMixin.dataframe() got an unexpected keyword argument 'use_container_width'`**: These errors indicate that features from newer Streamlit versions are being used with an older version (like 1.12.0). The code has been adjusted to use compatible alternatives (e.g., `st.experimental_rerun()`, removing `type` from button, removing `use_container_width` from dataframe).

## Limitations

- Works best with clearly structured, machine-printed invoices.
- OCR accuracy depends heavily on the quality and clarity of the scanned document.
- Complex layouts or handwritten invoices may result in lower accuracy.
=======
# Invoice-OCR-Custom-Field
>>>>>>> 23914deb889a1fb46c9416f90a73b48627b47e36
