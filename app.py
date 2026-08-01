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
            
            # Create a new Word Document
            doc = Document()
            
            for uploaded_file in uploaded_files:
                image = PIL.Image.open(uploaded_file)
                
                # 2. ASK THE AI FOR STRICT DATA
                prompt = """
                Extract the information from this court document image and output it as a JSON object with EXACTLY these keys:
                "thana" (for थाना)
                "apr_sankhya" (for अ०सं०)
                "vaad_sankhya" (for वाद सं०)
                "dhara" (for धारा)
                "vaadi" (for वादी का नाम व पता)
                "vivechak" (for विवेचक का नाम)
                "nirnay_date" (for निर्णय का दि०)
                "ghatna" (for घटना का संक्षिप्त विवरण)
                "adesh" (for न्यायालय के आदेश का विवरण)
                
                If any info is missing, use "-". Output ONLY valid JSON and nothing else.
                """
                
                try:
                    # Send to Gemini
                    response = client.models.generate_content(
                        model="gemini-1.5-flash", 
                        contents=[image, prompt]
                    )
                    
                    # Clean and load the data
                    raw_data = response.text.replace("```json", "").replace("```", "").strip()
                    data = json.loads(raw_data)
                    
                    # 3. BUILD THE TABLE WITH BORDERS IN WORD
                    doc.add_heading('न्यायालय डेली रिपोर्ट', level=1)
                    table = doc.add_table(rows=2, cols=9)
                    table.style = 'Table Grid' # This adds the black borders!
                    
                    # Set the Headers
                    headers = ['थाना', 'अ०सं०', 'वाद सं०', 'धारा', 'वादी का नाम व पता', 'विवेचक का नाम', 'निर्णय का दि०', 'घटना का संक्षिप्त विवरण', 'न्यायालय के आदेश का विवरण']
                    hdr_cells = table.rows[0].cells
                    for i, header in enumerate(headers):
                        hdr_cells[i].text = header
                    
                    # Fill the Data
                    row_cells = table.rows[1].cells
                    row_cells[0].text = data.get("thana", "-")
                    row_cells[1].text = data.get("apr_sankhya", "-")
                    row_cells[2].text = data.get("vaad_sankhya", "-")
                    row_cells[3].text = data.get("dhara", "-")
                    row_cells[4].text = data.get("vaadi", "-")
                    row_cells[5].text = data.get("vivechak", "-")
                    row_cells[6].text = data.get("nirnay_date", "-")
                    row_cells[7].text = data.get("ghatna", "-")
                    row_cells[8].text = data.get("adesh", "-")
                    
                    # Add a page break so the next photo goes on a new page
                    doc.add_page_break() 
                    
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: Check if the photo is clear.")
            
            # 4. CREATE THE DOWNLOAD BUTTON
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("✅ All documents processed successfully!")
            st.download_button(
                label="⬇️ Download Formatted Word Document",
                data=bio.getvalue(),
                file_name="Court_Daily_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
