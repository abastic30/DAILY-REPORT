import streamlit as st
from google import genai
import PIL.Image
import io
import json
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

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

# Helper function to build the EXACT vertical Word format
def add_case_to_word(doc, data):
    # --- 1. TOP HEADER SECTION ---
    header_table = doc.add_table(rows=2, cols=3)
    
    # Top Left
    p_left = header_table.cell(0, 0).paragraphs[0]
    p_left.add_run("P O- श्री ............................\nआरोप बनने का दि० ....................").bold = True
    
    # Top Center
    p_center = header_table.cell(0, 1).paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_center = p_center.add_run("न्यायालय डेली रिपोर्ट")
    run_center.bold = True
    run_center.font.size = Pt(16)
    
    # Top Right
    p_right = header_table.cell(0, 2).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.add_run("अभियोजक का नाम- श्री संजीव सिंह.\nदिनांक - .............................").bold = True
    
    # Court Name
    court_cell = header_table.cell(1, 0)
    court_cell.merge(header_table.cell(1, 1))
    court_cell.merge(header_table.cell(1, 2))
    p_court = court_cell.paragraphs[0]
    run_court = p_court.add_run("न्यायालय - ACJM - Kadipur सुल्तानपुर")
    run_court.bold = True
    run_court.font.size = Pt(14)
    
    doc.add_paragraph("") # Space before table
    
    # --- 2. MAIN DATA TABLE (4 Vertical Columns) ---
    table = doc.add_table(rows=8, cols=4)
    table.style = 'Table Grid'
    table.autofit = False
    
    # Set the widths (Total 7.5 inches for standard page)
    col_widths = [Inches(1.0), Inches(1.5), Inches(2.2), Inches(2.8)]
    
    # Fill Column 1: Labels
    labels = ['थाना', 'अ०सं०', 'वाद सं०', 'धारा', 'वादी का नाम व पता', 'विवेचक का नाम', 'निर्णय का दि०', 'अभियुक्त का नाम व पता']
    for i, label in enumerate(labels):
        table.cell(i, 0).text = label
        
    # Fill Column 2: Values
    table.cell(0, 1).text = str(data.get("thana", "-"))
    table.cell(1, 1).text = str(data.get("apr_sankhya", "-"))
    table.cell(2, 1).text = str(data.get("vaad_sankhya", "-"))
    table.cell(3, 1).text = str(data.get("dhara", "-"))
    table.cell(4, 1).text = str(data.get("vaadi", "-"))
    table.cell(5, 1).text = str(data.get("vivechak", "-"))
    table.cell(6, 1).text = str(data.get("nirnay_date", "-"))
    table.cell(7, 1).text = str(data.get("abhiyukt", "-"))
    
    # Fill Column 3: Ghatna (Merged block)
    table.cell(0, 2).text = "घटना का संक्षिप्त विवरण"
    table.cell(0, 2).paragraphs[0].runs[0].bold = True
    cell_ghatna_start = table.cell(1, 2)
    cell_ghatna_end = table.cell(7, 2)
    cell_ghatna_start.merge(cell_ghatna_end)
    cell_ghatna_start.text = str(data.get("ghatna", "-"))
    
    # Fill Column 4: Adesh (Merged block)
    table.cell(0, 3).text = "न्यायालय के आदेश का विवरण"
    table.cell(0, 3).paragraphs[0].runs[0].bold = True
    cell_adesh_start = table.cell(1, 3)
    cell_adesh_end = table.cell(7, 3)
    cell_adesh_start.merge(cell_adesh_end)
    cell_adesh_start.text = str(data.get("adesh", "-"))
    
    # Apply widths to all cells
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width
            
    doc.add_paragraph("\n") # Space before signatures
    
    # --- 3. BOTTOM FOOTER SECTION ---
    sig_table = doc.add_table(rows=1, cols=2)
    sig_left = sig_table.cell(0, 0).paragraphs[0]
    sig_left.add_run("का ० अभय राज सिंह\nनाम व हस्ताक्षर कोर्ट मोहर्रिर").bold = True
    
    sig_right = sig_table.cell(0, 1).paragraphs[0]
    sig_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig_right.add_run("हस्ताक्षर अभियोजक").bold = True
    
    doc.add_page_break()

# --- UI ELEMENTS ---
st.write("### 1. Choose Processing Mode")
mode = st.radio(
    "How should multiple photos be handled?", 
    ["📂 Combine all photos into ONE case report", 
     "📄 Process each photo as a SEPARATE case report"]
)

st.write("### 2. Add Manual Notes (Optional)")
user_notes = st.text_area("Type missing names, context, or corrections here:", placeholder="Example: The Vaadi is Rakesh. Case Date: 2024.")

st.write("### 3. Upload & Generate")
uploaded_files = st.file_uploader("Upload Document Photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Generate Word Document"):
        with st.spinner("Processing documents and generating your file..."):
            
            doc = Document()
            # Enforce Vertical/Portrait Mode
            section = doc.sections[0]
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.5)
            section.page_height = Inches(11.0)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            
            prompt = f"""
            Extract the information from this court document and output it as a JSON object with EXACTLY these 10 keys:
            "thana", "apr_sankhya", "vaad_sankhya", "dhara", "vaadi", "vivechak", "nirnay_date", "abhiyukt", "ghatna", "adesh".
            If any info is missing, use "-". Output ONLY valid JSON and nothing else.
            
            USER NOTES TO INCLUDE: {user_notes if user_notes else "None"}
            """
            
            success_count = 0
            
            if "Combine" in mode:
                images = [PIL.Image.open(f) for f in uploaded_files]
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
