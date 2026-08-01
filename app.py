import streamlit as st
from google import genai
from google.genai import types
import PIL.Image
import io
import json
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

# Set up the web app
st.set_page_config(page_title="Court Daily Report Generator", page_icon="⚖️", layout="wide")
st.title("📄 Court Daily Report Generator")

try:
    api_key = st.secrets["GEMINI_API_KEY"]
except KeyError:
    st.error("API Key not found. Please check Streamlit secrets.")
    st.stop()

client = genai.Client(api_key=api_key)

# Helper function to build the vertical Word format with COLORS
def add_case_to_word(doc, data):
    header_table = doc.add_table(rows=2, cols=3)
    
    p_left = header_table.cell(0, 0).paragraphs[0]
    p_left.add_run("PO- श्री ............................\nआरोप बनने का दि० ....................").bold = True
    
    p_center = header_table.cell(0, 1).paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_center = p_center.add_run("न्यायालय डेली रिपोर्ट")
    run_center.bold = True
    run_center.font.size = Pt(16)
    
    p_right = header_table.cell(0, 2).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.add_run("अभियोजक का नाम- श्री संजीव सिंह.\nदिनांक - .............................").bold = True
    
    court_cell = header_table.cell(1, 0)
    court_cell.merge(header_table.cell(1, 1))
    court_cell.merge(header_table.cell(1, 2))
    p_court = court_cell.paragraphs[0]
    run_court = p_court.add_run("न्यायालय - ACJM - Kadipur सुल्तानपुर")
    run_court.bold = True
    run_court.font.size = Pt(14)
    
    doc.add_paragraph("")
    
    table = doc.add_table(rows=8, cols=4)
    table.style = 'Table Grid'
    table.autofit = False
    
    col_widths = [Inches(1.0), Inches(1.5), Inches(2.2), Inches(2.8)]
    
    labels = ['थाना', 'अ०सं०', 'वाद सं०', 'धारा', 'वादी का नाम व पता', 'विवेचक का नाम', 'निर्णय का दि०', 'अभियुक्त का नाम व पता']
    for i, label in enumerate(labels):
        table.cell(i, 0).text = label
        
    table.cell(0, 1).text = str(data.get("thana", "-"))
    table.cell(1, 1).text = str(data.get("apr_sankhya", "-"))
    table.cell(2, 1).text = str(data.get("vaad_sankhya", "-"))
    table.cell(3, 1).text = str(data.get("dhara", "-"))
    table.cell(4, 1).text = str(data.get("vaadi", "-"))
    table.cell(5, 1).text = str(data.get("vivechak", "-"))
    table.cell(6, 1).text = str(data.get("nirnay_date", "-"))
    table.cell(7, 1).text = str(data.get("abhiyukt", "-"))
    
    # Fill Column 3: Ghatna (BLUE TEXT)
    table.cell(0, 2).text = "घटना का संक्षिप्त विवरण"
    table.cell(0, 2).paragraphs[0].runs[0].bold = True
    cell_ghatna_start = table.cell(1, 2)
    cell_ghatna_end = table.cell(7, 2)
    cell_ghatna_start.merge(cell_ghatna_end)
    cell_ghatna_start.text = "" 
    run_ghatna = cell_ghatna_start.paragraphs[0].add_run(str(data.get("ghatna", "-")))
    run_ghatna.font.color.rgb = RGBColor(0, 0, 255) 
    
    # Fill Column 4: Adesh (RED TEXT)
    table.cell(0, 3).text = "न्यायालय के आदेश का विवरण"
    table.cell(0, 3).paragraphs[0].runs[0].bold = True
    cell_adesh_start = table.cell(1, 3)
    cell_adesh_end = table.cell(7, 3)
    cell_adesh_start.merge(cell_adesh_end)
    cell_adesh_start.text = "" 
    run_adesh = cell_adesh_start.paragraphs[0].add_run(str(data.get("adesh", "-")))
    run_adesh.font.color.rgb = RGBColor(255, 0, 0) 
    
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width
            
    doc.add_paragraph("\n")
    
    sig_table = doc.add_table(rows=1, cols=2)
    sig_left = sig_table.cell(0, 0).paragraphs[0]
    sig_left.add_run("का ० अभय राज सिंह\nनाम व हस्ताक्षर कोर्ट मोहर्रिर").bold = True
    
    sig_right = sig_table.cell(0, 1).paragraphs[0]
    sig_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig_right.add_run("हस्ताक्षर अभियोजक").bold = True
    
    doc.add_page_break()

st.write("### 1. Choose Processing Mode")
mode = st.radio(
    "How should multiple photos be handled?", 
    ["📂 Combine all photos into ONE case report", 
     "📄 Process each photo as a SEPARATE case report"]
)

st.write("### 2. Manual Text Entry (Optional)")
with st.expander("📝 Tap here to manually type details (Overrides AI extraction)"):
    man_thana = st.text_input("थाना (Thana):")
    col1, col2 = st.columns(2)
    with col1:
        man_apr = st.text_input("अ०सं० (Crime No):")
        man_dhara = st.text_input("धारा (Sections):")
        man_vaadi = st.text_area("वादी का नाम व पता (Complainant Name/Address):", height=68)
        man_nirnay = st.text_input("निर्णय का दि० (Judgment Date):")
    with col2:
        man_vaad = st.text_input("वाद सं० (Case No):")
        man_vivechak = st.text_input("विवेचक का नाम (Vivechak Name):")
        man_abhiyukt = st.text_area("अभियुक्त का नाम व पता (Accused Name/Address):", height=68)
    man_ghatna = st.text_area("घटना का संक्षिप्त विवरण (Ghatna):")
    man_adesh = st.text_area("न्यायालय के आदेश का विवरण (Aadesh):")

st.write("### 3. Audio & Photo Upload")
st.info("🎙️ You can record your voice directly in the app, upload photos, or do both together!")

# NEW: Built-in Streamlit Audio Recorder
audio_value = st.audio_input("Record Voice Note (Dictate the case details)")

uploaded_files = st.file_uploader("Upload Document Photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# Generate if they provided EITHER photos OR an audio recording
if uploaded_files or audio_value:
    if st.button("Generate Word Document"):
        with st.spinner("Listening, reading, and generating your file..."):
            
            doc = Document()
            section = doc.sections[0]
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.5)
            section.page_height = Inches(11.0)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            
            # Package the audio if they recorded something
            audio_parts = []
            if audio_value is not None:
                audio_parts.append(
                    types.Part.from_bytes(
                        data=audio_value.getvalue(),
                        mime_type=audio_value.type
                    )
                )
            
            prompt = f"""
            Extract the information from the provided audio dictation and/or court document image. Output as a JSON object with EXACTLY these 10 keys:
            "thana", "apr_sankhya", "vaad_sankhya", "dhara", "vaadi", "vivechak", "nirnay_date", "abhiyukt", "ghatna", "adesh".
            
            PRIORITY RULE: If an audio file is provided, prioritize the spoken details over the image.
            
            CRITICAL INSTRUCTION FOR "ghatna" (घटना का संक्षिप्त विवरण):
            If generating from sections, use these EXACT legal templates:
            1. IF IPC 323, 504: Write EXACTLY: "अभियुक्तगण ने वादी के साथ गाली-गलौज की तथा मारपीट कर साधारण उपहति (चोट) कारित की। अपराध धारा [Insert Sections] के अंतर्गत दण्डनीय है।"
            2. IF Excise Act (60 or 63): Write EXACTLY: "अभियुक्त के पास से अवैध शराब बरामद हुई। अपराध धारा [Insert Sections] के अंतर्गत दण्डनीय है।"
            3. FOR ALL OTHER SECTIONS: Formulate a single, simple sentence describing the act, ending with "अपराध धारा [Insert Sections] के अंतर्गत दण्डनीय है।"
            
            IMPORTANT OVERRIDES (Use exactly as provided if not empty):
            Thana: {man_thana if man_thana else ""}
            Crime No: {man_apr if man_apr else ""}
            Case No: {man_vaad if man_vaad else ""}
            Section: {man_dhara if man_dhara else ""}
            Complainant: {man_vaadi if man_vaadi else ""}
            Investigator: {man_vivechak if man_vivechak else ""}
            Judgment Date: {man_nirnay if man_nirnay else ""}
            Accused: {man_abhiyukt if man_abhiyukt else ""}
            Ghatna: {man_ghatna if man_ghatna else ""}
            Adesh: {man_adesh if man_adesh else ""}
            
            If any info is missing, use "-". Output ONLY valid JSON.
            """
            
            success_count = 0
            
            try:
                # If they chose to combine everything, OR if they only provided audio and no photos
                if "Combine" in mode or not uploaded_files:
                    images = [PIL.Image.open(f) for f in uploaded_files] if uploaded_files else []
                    contents = audio_parts + images + [prompt] 
                    
                    response = client.models.generate_content(model="gemini-3.6-flash", contents=contents)
                    raw_text = response.text.strip()
                    if "```json" in raw_text:
                        raw_text = raw_text.split("```json")[1].split("```")[0].strip()
                    elif "```" in raw_text:
                        raw_text = raw_text.split("```")[1].split("```")[0].strip()
                        
                    data = json.loads(raw_text)
                    add_case_to_word(doc, data)
                    success_count += 1
                
                # If they want to process multiple separate cases (photos) individually
                else:
                    for uploaded_file in uploaded_files:
                        image = PIL.Image.open(uploaded_file)
                        # We attach the audio to EACH photo if they selected separate mode
                        contents = audio_parts + [image, prompt]
                        
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
                st.error(f"Error processing. Technical Details: {e}")
            
            if success_count > 0:
                bio = io.BytesIO()
                doc.save(bio)
                st.success(f"✅ Generated {success_count} report(s) successfully!")
                st.download_button(
                    label="⬇️ Download Formatted Word Document",
                    data=bio.getvalue(),
                    file_name="Court_Daily_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
