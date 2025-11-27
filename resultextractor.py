import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# --- Page Configuration ---
st.set_page_config(
    page_title="VTU Result Extractor",
    page_icon="🎓",
    layout="wide"
)

# --- Core Extraction Logic ---
def extract_data_from_pdf(uploaded_file):
    """
    Extracts student results from a single PDF file object.
    """
    extracted_data = []
    
    try:
        # pdfplumber.open can read directly from the uploaded file object
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                
                # 1. Extract Student Info (USN and Name)
                # Regex to find USN and Name
                usn_pattern = r"University Seat Number\s*[:,-]?\s*([0-9A-Z]+)"
                name_pattern = r"Student Name\s*[:,-]?\s*([A-Za-z\s]+)"
                
                usn_match = re.search(usn_pattern, text)
                name_match = re.search(name_pattern, text)
                
                usn = usn_match.group(1).strip() if usn_match else "Unknown"
                name = name_match.group(1).strip() if name_match else "Unknown"
                
                # 2. Extract Tables
                tables = page.extract_tables()
                
                for table in tables:
                    for row in table:
                        # Clean the row data (remove newlines inside cells)
                        clean_row = [str(item).replace('\n', ' ').strip() if item else '' for item in row]
                        
                        # Logic to identify result rows:
                        # 1. Row must have enough columns
                        # 2. First column must look like a subject code (e.g., BMATE201)
                        if len(clean_row) >= 6 and re.match(r'^[A-Z0-9]{5,}', clean_row[0]):
                            
                            sub_code = clean_row[0]
                            sub_name = clean_row[1]
                            internal = clean_row[2]
                            external = clean_row[3]
                            total = clean_row[4]
                            result = clean_row[5]

                            # FIX: Handle merged marks (e.g. Internal "25 40" and empty External)
                            if not external and ' ' in internal:
                                parts = internal.split()
                                if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                                    internal = parts[0]
                                    external = parts[1]

                            extracted_data.append({
                                "USN": usn,
                                "Student Name": name,
                                "Subject Code": sub_code,
                                "Subject Name": sub_name,
                                "Internal": internal,
                                "External": external,
                                "Total": total,
                                "Result": result,
                                "Source File": uploaded_file.name
                            })
    except Exception as e:
        st.error(f"Error processing {uploaded_file.name}: {e}")
        return []

    return extracted_data

def to_excel(df):
    """
    Converts dataframe to excel bytes for download.
    """
    output = io.BytesIO()
    # Use 'xlsxwriter' or 'openpyxl' engine
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Results')
    processed_data = output.getvalue()
    return processed_data

# --- UI Layout ---
st.title("🎓 VTU Result Extractor")
st.markdown("""
Upload your VTU Result PDFs below to convert them into a single Excel sheet.
**Instructions:**
1. Drag and drop one or multiple PDF files.
2. Check the data preview.
3. Click 'Download Excel'.
""")

# File Uploader
uploaded_files = st.file_uploader(
    "Upload PDF Files", 
    type="pdf", 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button(f"Process {len(uploaded_files)} Files"):
        all_results = []
        
        # Progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, pdf_file in enumerate(uploaded_files):
            # Update status
            status_text.text(f"Processing {pdf_file.name}...")
            
            # Extract
            data = extract_data_from_pdf(pdf_file)
            all_results.extend(data)
            
            # Update progress
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        status_text.text("Processing Complete!")
        
        if all_results:
            df = pd.DataFrame(all_results)
            
            st.success(f"Successfully extracted {len(df)} rows of data!")
            
            # Show Preview
            with st.expander("👁️ Preview Extracted Data", expanded=True):
                st.dataframe(df)
            
            # Excel Download Button
            excel_data = to_excel(df)
            st.download_button(
                label="📥 Download Excel File",
                data=excel_data,
                file_name="VTU_Bulk_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No result data found in the uploaded PDFs. Please check if the format matches.")