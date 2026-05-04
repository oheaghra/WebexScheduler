import streamlit as st
import pandas as pd
import json
import io
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- HELPER FUNCTIONS ---
def clean_housing(text):
    if not text or pd.isna(text): return ""
    return str(text).upper().replace("POD", "").strip()

def clean_webex(staff_name, staff_email):
    text = (str(staff_name) + " " + str(staff_email)).upper()
    mapping = {"ALPHA": "A", "APOD": "A", "BRAVO": "B", "BPOD": "B", "CHARLIE": "C", 
               "CPOD": "C", "EDWARD": "E", "EPOD": "E", "FOXTROT": "F", "FPOD": "F", 
               "INDIA": "I", "IPOD": "I", "CENTRAL": "CENTRAL"}
    for key, val in mapping.items():
        if key in text: return val
    return "Check File"

def format_inmate_info(name_str, dob_str):
    names = [n.strip() for n in str(name_str).split(';') if n.strip()]
    dobs = [d.strip() for d in str(dob_str).split(';') if d.strip()]
    lines = []
    for i in range(len(names)):
        n = names[i]
        d = dobs[i] if i < len(dobs) else "N/A"
        lines.append(f"{n.upper()} (DOB: {d})")
    return "\n".join(lines) if lines else "N/A"

# --- STREAMLIT WEB INTERFACE ---
st.set_page_config(page_title="Webex Scheduler", layout="wide")
st.title("📂 Webex Schedule Formatter")
st.write("Upload your .TSV file below to generate the Word document.")

uploaded_file = st.file_uploader("Choose a TSV file", type="tsv")

if uploaded_file is not None:
    # Read the TSV
    df = pd.read_csv(uploaded_file, sep='\t')
    
    # Process Word Doc in memory
    doc = Document()
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    try:
        date_obj = datetime.strptime(df.iloc[0]['Date Time'], '%m/%d/%Y %I:%M %p')
        header_date = date_obj.strftime('%m/%d/%Y')
    except:
        header_date = "Scheduled Date"

    doc.add_heading(f'Webex Schedule {header_date}', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    headers = ['Time', 'Inmate Information', 'Housing', 'Attorney Information', 'Webex']
    
    for i, h in enumerate(headers):
        run = table.rows[0].cells[i].paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(11)

    for _, row in df.iterrows():
        dt = datetime.strptime(row['Date Time'], '%m/%d/%Y %I:%M %p')
        time_str = (dt + timedelta(minutes=15)).strftime('%I:%M %p')
        custom = json.loads(row[' Custom Fields']) if pd.notna(row[' Custom Fields']) else {}
        
        row_cells = table.add_row().cells
        row_cells[0].paragraphs[0].add_run(time_str).font.size = Pt(11)
        row_cells[1].paragraphs[0].add_run(format_inmate_info(custom.get('INMATE NAME'), custom.get("Inmate's DOB"))).font.size = Pt(11)
        row_cells[2].paragraphs[0].add_run(clean_housing(custom.get('INMATE HOUSING LOCATION:'))).font.size = Pt(11)
        
        p_att = row_cells[3].paragraphs[0]
        p_att.add_run(f"{str(row['Customer Name']).upper()}\n").bold = True
        p_att.add_run(f"{row['Customer Phone']}\n").font.size = Pt(12)
        email_run = p_att.add_run(str(row['Customer Email']))
        email_run.font.size = Pt(12)
        email_run.font.name = 'Courier New'
        
        row_cells[4].paragraphs[0].add_run(clean_webex(row['Staff Name'], row['Staff Email'])).font.size = Pt(11)

    # Save to a "buffer" so the user can download it
    target_name = f"Webex_Schedule_{header_date.replace('/', '-')}.docx"
    bio = io.BytesIO()
    doc.save(bio)
    
    st.success(f"Successfully processed {header_date}!")
    st.download_button(
        label="📥 Download Formatted Word File",
        data=bio.getvalue(),
        file_name=target_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )