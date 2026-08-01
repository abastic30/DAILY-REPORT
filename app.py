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
st.set_page_config(page_title="Master Court Assistant", page_icon="⚖️", layout="wide")
st.title("⚖️ Master Court Assistant")

try:
    api_key = st.secrets["GEMINI_API_KEY"]
except KeyError:
    st.error("API Key not found. Please check Streamlit secrets.")
    st.stop()

client = genai.Client(api_key=api_key)

# ==========================================
# FUNCTION 1: COURT DAILY REPORT FORMAT (LANDSCAPE)
# ==========================================
def add_daily_report_to_word(doc, data, mohrir_name):
    # Force Landscape for Daily Report to match image_12.png format
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)

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
    court_cell.merge(header_table.cell(1, 2))
    run_court = court_cell.paragraphs[0].add_run("न्यायालय - ACJM - Kadipur सुल्तानपुर")
    run_court.bold = True
    run_court.font.size = Pt(14)
    doc.add_paragraph("")
    
    table = doc.add_table(rows=8, cols=4)
    table.style = 'Table Grid'
    table.autofit = False
    
    # Custom widths optimized for Landscape view
    col_widths = [Inches(1.2), Inches(2.0), Inches(3.2), Inches(3.6)]
    
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
    
    table.cell(0, 2).text = "घटना का संक्षिप्त विवरण"
    table.cell(0, 2).paragraphs[0].runs[0].bold = True
    c_ghatna = table.cell(1, 2)
    c_ghatna.merge(table.cell(7, 2))
    c_ghatna.text = "" 
    c_ghatna.paragraphs[0].add_run(str(data.get("ghatna", "-"))).font.color.rgb = RGBColor(0, 0, 255) 
    
    table.cell(0, 3).text = "न्यायालय के आदेश का विवरण"
    table.cell(0, 3).paragraphs[0].runs[0].bold = True
    c_adesh = table.cell(1, 3)
    c_adesh.merge(table.cell(7, 3))
    c_adesh.text = "" 
    c_adesh.paragraphs[0].add_run(str(data.get("adesh", "-"))).font.color.rgb = RGBColor(255, 0, 0) 
    
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width
            
    doc.add_paragraph("\n")
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.cell(0, 0).paragraphs[0].add_run(f"{mohrir_name}\nनाम व हस्ताक्षर कोर्ट मोहर्रिर").bold = True
    sig_table.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig_table.cell(0, 1).paragraphs[0].add_run("हस्ताक्षर अभियोजक").bold = True
    doc.add_page_break()

# ==========================================
# FUNCTION 2: SOOCHNA (NOTICE) FORMAT
# ==========================================
def add_soochna_to_word(doc, data, action_type, mohrir_name):
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_head.add_run("सूचना\n")
    run_title.bold = True
    run_title.underline = True
    run_title.font.size = Pt(18)
    
    court_name = str(data.get("court_name", "JM-JD"))
    run_court = p_head.add_run(f"न्यायालय {court_name}\nSULTANPUR")
    run_court.bold = True
    run_court.font.size = Pt(14)
    doc.add_paragraph("")

    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(3.0)
    table.columns[1].width = Inches(4.0)
    
    accused = str(data.get("accused", "...................................."))
    table.cell(0, 0).text = f"नाम पता अभियुक्त-\n{accused}"
    
    crime = str(data.get("crime_no", "............."))
    dhara = str(data.get("dhara", "............."))
    thana = str(data.get("thana", "............."))
    table.cell(0, 1).text = f"मु ०अ ०सं०- {crime}\nधारा - {dhara}\n\nथाना- {thana}\nजिला- सुल्तानपुर"
    
    thana_sho = str(data.get("thana", "............."))
    date_app = str(data.get("date", ".................."))
    
    p_body = doc.add_paragraph()
    p_body.add_run(f"\nसेवा में,\n      श्रीमान SHO\n      थाना {thana_sho}, जिला सुल्तानपुर\nमहोदय,\n")
    action_text = "जमानत" if action_type == "जमानत (Bail)" else "NBW RECALL"
    p_body.add_run(f"      निवेदन के साथ अवगत कराना है कि मु ०अ ०सं० व धारा उपरोक्त में अभि० उपरोक्त दिनांक {date_app} को न्यायालय के समक्ष उपस्थित होकर {action_text} करा लिया है।\n{action_text} सूचना आवश्यक कार्यवाही हेतु प्रेषित है।")
    
    p_foot = doc.add_paragraph(f"\n\n{mohrir_name}\nको०मो०")
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in p_foot.runs:
        run.bold = True
    doc.add_page_break()

# ==========================================
# FUNCTION 3: WITNESS/SUMMONS REPORT FORMAT
# ==========================================
def add_witness_report_to_word(doc, data, mohrir_name):
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    court_name = str(data.get("court_name", "JM (JD) OUTLINE COURT कादीपुर सुल्तानपुर"))
    run_title = p_head.add_run(f"न्यायालय {court_name}\n")
    run_title.bold = True
    run_title.font.size = Pt(16)
    
    report_date = str(data.get("report_date", ".................."))
    run_sub = p_head.add_run(f"साक्षिगणों को निर्गत समन व उनकी उपस्थिति के संबंध में विवरण दिनांक {report_date}")
    run_sub.font.size = Pt(14)
    
    table = doc.add_table(rows=2, cols=8)
    table.style = 'Table Grid'
    
    headers = [
        "जनपद", "साक्ष्य हेतु\nनियत वाद", "निर्गत कुल\nसम्मन", 
        "तामीला\nहोकर\nप्राप्त\nसम्मन", "निर्गत कुल\nवारंट", 
        "तामीला\nहोकर प्राप्त\nवारंट", "उपस्थित आए\nसाक्षियों की\nसंख्या", 
        "परीक्षित\nसाक्षियों की\nसंख्या"
    ]
    
    for i, head_text in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = head_text
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].runs[0].bold = True
        
    keys = ["janpad", "niyat_vaad", "nirgat_samman", "tamila_samman", "nirgat_warrant", "tamila_warrant", "upasthit_sakshi", "parikshit_sakshi"]
    for i, key in enumerate(keys):
        cell = table.cell(1, i)
        cell.text = str(data.get(key, "00"))
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    doc.add_paragraph("")
    vc_count = str(data.get("vc_count", "0"))
    p_foot = doc.add_paragraph(f"VC से होने वाली गवाही सूचना - {vc_count}")
    p_foot.runs[0].font.size = Pt(14)
    doc.add_page_break()

# ==========================================
# MAIN APP UI
# ==========================================
st.markdown("### ⚙️ Global Settings")
global_mohrir = st.text_input("✍️ Court Mohrir Name (कोर्ट मोहर्रिर):", value="का ० अभय राज सिंह")
st.markdown("---")

doc_type = st.selectbox("📑 Select Document Type to Generate:", 
                        ["Court Daily Report (डेली रिपोर्ट)", 
                         "Soochna (सूचना - Jamanat / NBW Recall)",
                         "Witness/Summons Report (साक्षी/समन विवरण)"])
st.markdown("---")

# ----------------- UI: DAILY REPORT -----------------
if doc_type == "Court Daily Report (डेली रिपोर्ट)":
    mode = st.radio("How should multiple photos be handled?", ["📂 Combine all photos into ONE report", "📄 Process SEPARATE reports"])
    uploaded_files = st.file_uploader("Upload Photos (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    if uploaded_files:
        preview_cols = st.columns(len(uploaded_files) if len(uploaded_files) < 4 else 4)
        for i, file in enumerate(uploaded_files):
            preview_cols[i % 4].image(file, caption=f"Page {i+1}", use_container_width=True)
            
    audio_file = st.audio_input("🎙️ Record Dictation (Overrides Photo)")
    
    with st.expander("📝 Manual Text Overrides"):
        head_col1, head_col2, head_col3 = st.columns(3)
        man_po = head_col1.text_input("PO- श्री (Judge):")
        man_abhiyojak = head_col2.text_input("अभियोजक (Prosecutor):")
        man_date = head_col3.text_input("दिनांक (Date):")
        man_thana = st.text_input("थाना (Thana):")
        col1, col2 = st.columns(2)
        man_apr = col1.text_input("अ०सं० (Crime No):")
        man_dhara = col1.text_input("धारा (Sections):")
        man_vaadi = col1.text_area("वादी का नाम (Complainant):", height=68)
        man_nirnay = col1.text_input("निर्णय का दि०:")
        man_vaad = col2.text_input("वाद सं० (Case No):")
        man_vivechak = col2.text_input("विवेचक का नाम:")
        man_abhiyukt = col2.text_area("अभियुक्त का नाम (Accused):", height=68)
        man_ghatna = st.text_area("घटना का संक्षिप्त विवरण (Ghatna):")
        man_adesh = st.text_area("न्यायालय के आदेश का विवरण (Aadesh):")

    if (uploaded_files or audio_file) and st.button("Generate Daily Report", type="primary"):
        with st.spinner("Processing..."):
            doc = Document()
            prompt = f"""
            Extract information into JSON with keys: "po_name", "abhiyojak_name", "report_date", "thana", "apr_sankhya", "vaad_sankhya", "dhara", "vaadi", "vivechak", "nirnay_date", "abhiyukt", "ghatna", "adesh".
            CRITICAL GHATNA RULES: If IPC 323, 504 -> "अभियुक्तगण ने वादी के साथ गाली-गलौज की तथा मारपीट कर साधारण उपहति (चोट) कारित की। अपराध धारा [Sections] के अंतर्गत दण्डनीय है।" If 60/63 Excise -> "अभियुक्त के पास से अवैध शराब बरामद हुई। अपराध धारा [Sections] के अंतर्गत दण्डनीय है।"
            OVERRIDES: PO:{man_po} Abhiyojak:{man_abhiyojak} Date:{man_date} Thana:{man_thana} Crime:{man_apr} Case:{man_vaad} Dhara:{man_dhara} Vaadi:{man_vaadi} Vivechak:{man_vivechak} NirnayDate:{man_nirnay} Accused:{man_abhiyukt} Ghatna:{man_ghatna} Adesh:{man_adesh}. Output ONLY JSON.
            """
            audio_part = types.Part.from_bytes(data=audio_file.getvalue(), mime_type=audio_file.type) if audio_file else None
            try:
                if "Combine" in mode and uploaded_files:
                    contents = [PIL.Image.open(f) for f in uploaded_files]
                    if audio_part: contents.append(audio_part)
                    contents.append(prompt)
                    res = client.models.generate_content(model="gemini-3.6-flash", contents=contents).text.strip()
                    add_daily_report_to_word(doc, json.loads(res.replace("```json","").replace("```","").strip()), global_mohrir)
                elif uploaded_files:
                    for f in uploaded_files:
                        contents = [PIL.Image.open(f), audio_part, prompt] if audio_part else [PIL.Image.open(f), prompt]
                        res = client.models.generate_content(model="gemini-3.6-flash", contents=contents).text.strip()
                        add_daily_report_to_word(doc, json.loads(res.replace("```json","").replace("```","").strip()), global_mohrir)
                elif audio_file:
                    res = client.models.generate_content(model="gemini-3.6-flash", contents=[audio_part, prompt]).text.strip()
                    add_daily_report_to_word(doc, json.loads(res.replace("```json","").replace("```","").strip()), global_mohrir)
                    
                bio = io.BytesIO()
                doc.save(bio)
                st.success("✅ Generated in Landscape Mode successfully!")
                st.download_button("⬇️ Download Daily Report", data=bio.getvalue(), file_name="Court_Daily_Report.docx")
            except Exception as e: st.error(f"Error: {e}")

# ----------------- UI: SOOCHNA -----------------
elif doc_type == "Soochna (सूचना - Jamanat / NBW Recall)":
    action_choice = st.radio("📌 Select Notice Type:", ["जमानत (Bail)", "NBW RECALL"])
    audio_file_soochna = st.audio_input("🎙️ Record Dictation")
    uploaded_files_soochna = st.file_uploader("Upload FIR/Photos (Optional)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    
    with st.expander("📝 Manual Text Overrides"):
        s_col1, s_col2 = st.columns(2)
        s_court = s_col1.text_input("न्यायालय (Court Name):")
        s_accused = s_col1.text_area("नाम पता अभियुक्त (Accused):")
        s_date = s_col1.text_input("उपस्थिति दिनांक (Date):")
        s_crime = s_col2.text_input("मु०अ०सं० (Crime No):")
        s_dhara = s_col2.text_input("धारा (Sections):")
        s_thana = s_col2.text_input("थाना (Thana):")

    if (uploaded_files_soochna or audio_file_soochna or s_court or s_accused) and st.button("Generate Soochna Document", type="primary"):
        with st.spinner("Processing..."):
            doc = Document()
            prompt = f"""
            Extract details for 'Soochna' into JSON with EXACTLY these keys: "court_name", "accused", "crime_no", "dhara", "thana", "date".
            OVERRIDES (Use exactly if provided): Court:{s_court}, Accused:{s_accused}, CrimeNo:{s_crime}, Dhara:{s_dhara}, Thana:{s_thana}, Date:{s_date}. Output ONLY JSON.
            """
            audio_part = types.Part.from_bytes(data=audio_file_soochna.getvalue(), mime_type=audio_file_soochna.type) if audio_file_soochna else None
            contents = []
            if uploaded_files_soochna: contents.extend([PIL.Image.open(f) for f in uploaded_files_soochna])
            if audio_part: contents.append(audio_part)
            contents.append(prompt)
            
            try:
                res = client.models.generate_content(model="gemini-3.6-flash", contents=contents).text.strip()
                add_soochna_to_word(doc, json.loads(res.replace("```json","").replace("```","").strip()), action_choice, global_mohrir)
                bio = io.BytesIO()
                doc.save(bio)
                st.success(f"✅ {action_choice} Soochna generated!")
                st.download_button("⬇️ Download Soochna.docx", data=bio.getvalue(), file_name="Soochna_Notice.docx")
            except Exception as e: st.error(f"Error: {e}")

# ----------------- UI: WITNESS/SUMMONS REPORT -----------------
elif doc_type == "Witness/Summons Report (साक्षी/समन विवरण)":
    audio_file_wit = st.audio_input("🎙️ Record Dictation")
    uploaded_files_wit = st.file_uploader("Upload Table Photo (Optional)", type=["jpg", "jpeg", "png"])
    
    with st.expander("📝 Manual Number Entry (Fill the Table)"):
        st.markdown("**Header Details**")
        w_court = st.text_input("न्यायालय (Court Name):", placeholder="e.g., JM (JD) OUTLINE COURT कादीपुर सुल्तानपुर")
        w_date = st.text_input("दिनांक (Date):", placeholder="e.g., 25/07/2026")
        
        st.markdown("**Table Data (Numbers)**")
        w_col1, w_col2, w_col3, w_col4 = st.columns(4)
        w_janpad = w_col1.text_input("जनपद (District):", value="सुल्तानपुर")
        w_1 = w_col2.text_input("साक्ष्य हेतु नियत वाद:", placeholder="00")
        w_2 = w_col3.text_input("निर्गत कुल सम्मन:", placeholder="00")
        w_3 = w_col4.text_input("तामीला होकर प्राप्त सम्मन:", placeholder="00")
        
        w_col5, w_col6, w_col7, w_col8 = st.columns(4)
        w_4 = w_col5.text_input("निर्गत कुल वारंट:", placeholder="00")
        w_5 = w_col6.text_input("तामीला होकर प्राप्त वारंट:", placeholder="00")
        w_6 = w_col7.text_input("उपस्थित साक्षियों की संख्या:", placeholder="00")
        w_7 = w_col8.text_input("परीक्षित साक्षियों की संख्या:", placeholder="00")
        
        st.markdown("**Footer Details**")
        w_vc = st.text_input("VC से होने वाली गवाही सूचना (VC Count):", placeholder="0")

    if (uploaded_files_wit or audio_file_wit or w_date or w_1) and st.button("Generate Witness Report", type="primary"):
        with st.spinner("Processing..."):
            doc = Document() 
            prompt = f"""
            Extract details into JSON with EXACTLY these keys: "court_name", "report_date", "janpad", "niyat_vaad", "nirgat_samman", "tamila_samman", "nirgat_warrant", "tamila_warrant", "upasthit_sakshi", "parikshit_sakshi", "vc_count".
            OVERRIDES (Use exactly if provided): Court:{w_court}, Date:{w_date}, Janpad:{w_janpad}, NiyatVaad:{w_1}, NirgatSamman:{w_2}, TamilaSamman:{w_3}, NirgatWarrant:{w_4}, TamilaWarrant:{w_5}, Upasthit:{w_6}, Parikshit:{w_7}, VC:{w_vc}. Output ONLY JSON.
            """
            audio_part = types.Part.from_bytes(data=audio_file_wit.getvalue(), mime_type=audio_file_wit.type) if audio_file_wit else None
            contents = []
            if uploaded_files_wit: contents.append(PIL.Image.open(uploaded_files_wit))
            if audio_part: contents.append(audio_part)
            contents.append(prompt)
            
            try:
                res = client.models.generate_content(model="gemini-3.6-flash", contents=contents).text.strip()
                add_witness_report_to_word(doc, json.loads(res.replace("```json","").replace("```","").strip()), global_mohrir)
                bio = io.BytesIO()
                doc.save(bio)
                st.success("✅ Witness Report generated in Landscape Mode!")
                st.download_button("⬇️ Download Witness Report.docx", data=bio.getvalue(), file_name="Witness_Report.docx")
            except Exception as e: st.error(f"Error: {e}")
