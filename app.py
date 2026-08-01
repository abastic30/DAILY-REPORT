import streamlit as st
from google import genai
import PIL.Image
import io
import json
from docx import Document

# Set up the web app
st.set_page_config(page_title="Court Daily Report Generator", page_icon="⚖️", layout="centered")
st.title("📄 Court Daily Report Generator")

# Securely get API key
try:
    api_key = st.secrets["GEMINI_API_KEY"]
except KeyError:
    st.error("API Key not found. Please check Streamlit secrets.")
    st.stop()

client = genai.Client(api_key=api_key)

# 1. ALLOW MULTIPLE FILE UPLOADS
uploaded_files = st.file_uploader("Upload Document Photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Generate Word Document"):
        with st.spinner("Processing documents and generating your file..."):
            
            doc = Document()
            success_count = 0
            
            for uploaded_file in uploaded_files:
                image = PIL.Image.open(uploaded_file)
                
                # 2. STRICTER PROMPT
                prompt = """
                Extract the information from this court document image and output it as a JSON object with EXACTLY these keys:
                "thana", "apr_sankhya", "vaad_sankhya", "dhara", "vaadi", "vivechak", "nirnay_date", "ghatna", "adesh".
                If any info is missing, use "-". Output ONLY valid JSON and nothing else. No markdown, no introduction.
                """
                
                try:
                    # Send to Gemini
                    response = client.models.generate_content(
                        model="gemini-3.6-flash", 
                        contents=[image, prompt]
                    )
                    
                    # 3. SMARTER JSON CLEANUP
                    raw_text = response.text.strip()
                    
                    # Strip out markdown code blocks if the AI adds them
                    if "```json" in raw_text:
                        raw_text = raw_text.split("```json")[1].split("```")[0].strip()
                    elif "```" in raw_text:
                        raw_text = raw_text.split("```")[1].split("```")[0].strip()
                        
                    data = json.loads(raw_text)
                    
                    # 4. BUILD THE TABLE WITH BORDERS IN WORD
                    doc.add_heading('न्यायालय डेली रिपोर्ट', level=1)
                    table = doc.add_table(rows=2, cols=9)
                    table.style = 'Table Grid'
                    
                    # Set the Headers
                    headers = ['थाना', 'अ०सं०', 'वाद सं०', 'धारा', 'वादी का नाम व पता', 'विवेचक का नाम', 'निर्णय का दि०', 'घटना का संक्षिप्त विवरण', 'न्यायालय के आदेश का विवरण']
                    hdr_cells = table.rows[0].cells
                    for i, header in enumerate(headers):
                        hdr_cells[i].text = header
                    
                    # Fill the Data
                    row_cells = table.rows[1].cells
                    row_cells[0].text = str(data.get("thana", "-"))
                    row_cells[1].text = str(data.get("apr_sankhya", "-"))
                    row_cells[2].text = str(data.get("vaad_sankhya", "-"))
                    row_cells[3].text = str(data.get("dhara", "-"))
                    row_cells[4].text = str(data.get("vaadi", "-"))
                    row_cells[5].text = str(data.get("vivechak", "-"))
                    row_cells[6].text = str(data.get("nirnay_date", "-"))
                    row_cells[7].text = str(data.get("ghatna", "-"))
                    row_cells[8].text = str(data.get("adesh", "-"))
                    
                    doc.add_page_break() 
                    success_count += 1
                    
                except Exception as e:
                    # THIS WILL NOW SHOW EXACTLY WHY IT FAILED
                    st.error(f"Error processing {uploaded_file.name}. Technical Details: {e}")
                    if 'response' in locals() and hasattr(response, 'text'):
                        st.warning(f"What the AI tried to send back: {response.text}")
            
            # 5. CREATE DOWNLOAD BUTTON ONLY IF AT LEAST ONE SUCCEEDED
            if success_count > 0:
                bio = io.BytesIO()
                doc.save(bio)
                
                st.success(f"✅ Successfully processed {success_count} document(s)!")
                st.download_button(
                    label="⬇️ Download Formatted Word Document",
                    data=bio.getvalue(),
                    file_name="Court_Daily_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
