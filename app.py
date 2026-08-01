import streamlit as st
from google import genai
from google.genai import types  # Added to handle audio files
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

# Helper function to build the vertical Word format
def add_case_to_word(doc, data):
    header_table = doc.add_table(rows=2, cols=3)
    
    po_val = str(data.get("po_name", ""))
    po_text = po_val if po_val and po_val != "-" else "............................"
    
    abhi_val = str(data.get("abhiyojak_name", ""))
    abhi_text = abhi_val if abhi_val and abhi_val != "-" else "श्री संजीव सिंह"
    
    date_val = str(data.get("report_date", ""))
    date_text = date_val if date_val and date_val != "-" else "............................."
    
    p_left = header_table.cell(0, 0).paragraphs[0]
    p_left.add_run(f"PO- श्री {po_text}\nआरोप बनने का दि० ....................").bold = True
    
    p_center = header_table.cell(0, 1).paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_center = p_center.add_run("न्यायालय डेली रिपोर्ट")
    run_center.bold = True
    run_center.font.size = Pt(16)
    
    p_right = header_table.cell(0, 2).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.add_run(f"अभियोजक का नाम- {abhi_text}.\nदिनांक - {date_text}").bold = True
    
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

st.write("### 2. Upload Document Photos")
uploaded_files = st.file_uploader("Upload Photos (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# --- NEW: IMAGE PREVIEW GRID ---
if uploaded_files:
    st.markdown("🖼️ **Uploaded Images Preview:**")
    # Creates a neat row of image thumbnails
    preview_cols = st.columns(len(uploaded_files) if len(uploaded_files) < 4 else 4)
    for i, file in enumerate(uploaded_files):
        with preview_cols[i % 4]:
            st.image(file, caption=f"Page {i+1}", use_container_width=True)

st.write("### 3. Voice Dictation (Audio)")
st.info("🎙️ Tap the microphone below to dictate the case details. The AI will listen to you and prioritize your voice!")
# --- NEW: NATIVE AUDIO MICROPHONE INPUT ---
audio_file = st.audio_input("Record Dictation (Optional)")

st.write("### 4. Manual Text Overrides (Optional)")
with st.expander("📝 Tap here to manually type details (Overrides both Photo and Audio)"):
    st.markdown("**Header Details**")
    head_col1, head_col2, head_col3 = st.columns(3)
    with head_col1:
        man_po = st.text_input("PO- श्री (Judge Name):")
    with head_col2:
        man_abhiyojak = st.text_input("अभियोजक (Prosecutor):")
    with head_col3:
        man_date = st.text_input("दिनांक (Date):")
        
    st.markdown("---")
    st.markdown("**Case Details**")
    man_thana = st.text_input("थाना (Thana):")
    
    col1, col2 = st.columns(2)
    with col1:
        man_apr = st.text_input("अ०सं० (Crime No):")
        man_dhara = st.text_input("धारा (Sections):")
        man_vaadi = st.text_area("वादी का नाम व पता (Complainant):", height=68)
        man_nirnay = st.text_input("निर्णय का दि० (Judgment Date):")
        
    with col2:
        man_vaad = st.text_input("वाद सं० (Case No):")
        man_vivechak = st.text_input("विवेचक का नाम (Vivechak):")
        man_abhiyukt = st.text_area("अभियुक्त का नाम व पता (Accused):", height=68)
        
    man_ghatna = st.text_area("घटना का संक्षिप्त विवरण (Ghatna - Sankshipt Vivaran):")
    man_adesh = st.text_area("न्यायालय के आदेश का विवरण (Aadesh):")

st.write("### 5. Generate Report")
if uploaded_files or audio_file:
    if st.button("Generate Word Document", type="primary"):
        with st.spinner("Reading documents and listening to audio... Please wait..."):
            
            doc = Document()
            section = doc.sections[0]
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.5)
            section.page_height = Inches(11.0)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            
            prompt = f"""
            Extract the information from the provided court document images AND/OR the provided audio dictation. 
            Output it as a JSON object with EXACTLY these 13 keys:
            "po_name", "abhiyojak_name", "report_date", "thana", "apr_sankhya", "vaad_sankhya", "dhara", "vaadi", "vivechak", "nirnay_date", "abhiyukt", "ghatna", "adesh".
            
            AUDIO PRIORITY RULE: If an audio dictation was provided, prioritize the information spoken in the audio. Use the images to fill in any gaps the audio missed.
            
            CRITICAL INSTRUCTION FOR "ghatna" (घटना का संक्षिप्त विवरण):
            If generating from police information, read the "dhara" (Sections) and use these EXACT legal templates:
            1. IF IPC 323, 504: Write EXACTLY: "अभियुक्तगण ने वादी के साथ गाली-गलौज की तथा मारपीट कर साधारण उपहति (चोट) कारित की। अपराध धारा [Insert Sections] के अंतर्गत दण्डनीय है।"
            2. IF Excise Act (60 or 63): Write EXACTLY: "अभियुक्त के पास से अवैध शराब बरामद हुई। अपराध धारा [Insert Sections] के अंतर्गत दण्डनीय है।"
            3. FOR ALL OTHER SECTIONS: Formulate a single, simple sentence describing the act, ending with "अपराध धारा [Insert Sections] के अंतर्गत दण्डनीय है।"
            
            IMPORTANT TEXT OVERRIDES: If any of the following fields have data, YOU MUST USE IT EXACTLY AS PROVIDED instead of extracting it from the image or audio:
            PO Name: {man_po if man_po else "Extract"}
            Prosecutor: {man_abhiyojak if man_abhiyojak else "Extract from image/audio. Default to 'श्री संजीव सिंह' if none found."}
            Report Date: {man_date if man_date else "Extract"}
            Thana: {man_thana if man_thana else "Extract"}
            Crime No: {man_apr if man_apr else "Extract"}
            Case No: {man_vaad if man_vaad else "Extract"}
            Section: {man_dhara if man_dhara else "Extract"}
            Complainant: {man_vaadi if man_vaadi else "Extract"}
            Investigator: {man_vivechak if man_vivechak else "Extract"}
            Judgment Date: {man_nirnay if man_nirnay else "Extract"}
            Accused: {man_abhiyukt if man_abhiyukt else "Extract"}
            Ghatna: {man_ghatna if man_ghatna else "Follow the CRITICAL INSTRUCTION above"}
            Adesh: {man_adesh if man_adesh else "Extract"}
            
            If any info is missing, use "-". Output ONLY valid JSON.
            """
            
            # --- AUDIO PROCESSING LOGIC ---
            audio_part = None
            if audio_file:
                audio_part = types.Part.from_bytes(data=audio_file.getvalue(), mime_type=audio_file.type)
            
            success_count = 0
            
            # Scenario 1: Combine Images
            if "Combine" in mode and uploaded_files:
                images = [PIL.Image.open(f) for f in uploaded_files]
                contents = images
                if audio_part:
                    contents.append(audio_part)
                contents.append(prompt)
                
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
            
            # Scenario 2: Separate Images
            elif "SEPARATE" in mode and uploaded_files:
                for uploaded_file in uploaded_files:
                    image = PIL.Image.open(uploaded_file)
                    contents = [image]
                    if audio_part:
                        contents.append(audio_part) # Audio applies to each if separate mode is chosen
                    contents.append(prompt)
                    
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
                        st.error(f"Error processing {uploaded_file.name}. Details: {e}")
            
            # Scenario 3: Audio ONLY (No Images uploaded)
            elif audio_file and not uploaded_files:
                 try:
                    contents = [audio_part, prompt]
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
                    st.error(f"Error processing audio. Details: {e}")
            
            if success_count > 0:
                bio = io.BytesIO()
                doc.save(bio)
                st.success(f"✅ Document generated successfully!")
                st.download_button(
                    label="⬇️ Download Formatted Word Document",
                    data=bio.getvalue(),
                    file_name="Court_Daily_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
