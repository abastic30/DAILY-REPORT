import streamlit as st
from google import genai
import PIL.Image
import io
import json
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_SECTION

# Set up the web app
st.set_page_config(page_title="Court Daily Report Generator", page_icon="⚖️", layout="wide")
st.title("📄 Court Daily Report Generator")

# Securely get API key
try:
    api_key = st.secrets["GEMINI_API_KEY"]
except KeyError:
    st.error("API Key not found. Please check Streamlit secrets.")
    st.stop()

client = genai.Client(api_key=api_key)

# Helper function to build the Word table (keeps code clean)
def add_case_to_word(doc, data):
    doc.add_heading('न्यायालय डेली रिपोर्ट', level=1)
    table = doc.add_table(rows=2, cols=9)
    table.style = 'Table Grid'
    table.autofit = False
    
    col_widths = [Inches(0.8), Inches(0.7), Inches(0.7), Inches(0.7), Inches(1.1), Inches(0.9), Inches(0.8), Inches(1.8), Inches(2.5)]
    headers = ['थाना', 'अ०सं०', 'वाद सं०', 'धारा', 'वादी का नाम व पता', 'विवेचक का नाम', 'निर्णय का दि०', 'घटना का संक्षिप्त विवरण', 'न्यायालय के आदेश का विवरण']
    
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        
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
    
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width
    doc.add_page_break()

# --- NEW UI ELEMENTS ---
st.write("### 1. Choose Processing Mode")
mode = st.radio(
    "How should multiple photos be handled?", 
    ["📂 Combine all photos into ONE case report (e.g., a 3-page document)", 
     "📄 Process each photo as a SEPARATE case report"]
)

st.write("### 2. Add Manual Notes (Optional)")
user_notes = st.text_area("Type missing names, context, or corrections here to help the AI:", placeholder="Example: The Vaadi is Rakesh Kumar. The date is 14/08/2023.")

st.write("### 3. Upload & Generate")
uploaded_files = st.file_uploader("Upload Document Photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Generate Word Document"):
        with st.spinner("Processing documents and generating your file..."):
            
            doc = Document()
            section = doc.sections[0]
            section.orientation = WD_SECTION.LANDSCAPE
            new_width, new_height = section.page_height, section.page_width
            section.page_width = new_width
            section.page_height = new_height
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            
            prompt = f"""
            Extract the information from this court document and output it as a JSON object with EXACTLY these keys:
            "thana", "apr_sankhya", "vaad_sankhya", "dhara", "vaadi", "vivechak", "nirnay_date", "ghatna", "adesh".
            If any info is missing, use "-". Output ONLY valid JSON and nothing else. No markdown.
            
            IMPORTANT USER NOTES TO INCLUDE: {user_notes if user_notes else "None"}
            """
            
            success_count = 0
            
            # --- LOGIC FOR COMBINING PHOTOS ---
            if "Combine" in mode:
                images = [PIL.Image.open(f) for f in uploaded_files]
                # Send all images in a single request to the AI
                contents = images + [prompt] 
                
                try:
                    response = client.models.generate_content(model="gemini-3.6-flash", contents=contents)
                    raw_text = response.text.strip()
                    if "```json" in raw_text:
                        raw_text = raw_text.split("```json")[1].split("```")[0].strip()
                    elif "```" in raw_text:
                        raw_text = raw_text.split("```")[1].split("```")[0].strip()
                        
                    data = json.loads(raw_text)
                    add_case_to_word(doc, data)
                    success_count += 1
                except Exception as e:
                    st.error(f"Error processing combined photos. Details: {e}")

            # --- LOGIC FOR SEPARATE PHOTOS ---
            else:
                for uploaded_file in uploaded_files:
                    image = PIL.Image.open(uploaded_file)
                    try:
                        response = client.models.generate_content(model="gemini-3.6-flash", contents=[image, prompt])
                        raw_text = response.text.strip()
                        if "```json" in raw_text:
                            raw_text = raw_text.split("```json")[1].split("```")[0].strip()
                        elif "```" in raw_text:
                            raw_text = raw_text.split("```")[1].split("```")[0].strip()
                            
                        data = json.loads(raw_text)
                        add_case_to_word(doc, data)
                        success_count += 1
                    except Exception as e:
                        st.error(f"Error processing {uploaded_file.name}. Details: {e}")
            
            if success_count > 0:
                bio = io.BytesIO()
                doc.save(bio)
                st.success("✅ Document processed successfully!")
                st.download_button(
                    label="⬇️ Download Formatted Word Document",
                    data=bio.getvalue(),
                    file_name="Court_Daily_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
