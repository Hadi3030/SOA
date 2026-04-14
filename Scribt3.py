import streamlit as st
import pandas as pd
import io

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

def set_cell_shading(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)
    
st.set_page_config(page_title="SOA Report ALL Broker", layout="wide")

# ===============================
# HEADER
# ===============================
col1, col2 = st.columns([1, 6])

with col1:
    st.image("askrindo.jpg", width=120)

with col2:
    st.title("📑 SOA Report Generator - All Broker")

# ===============================
# UPLOAD
# ===============================
file = st.file_uploader("Upload File SOA", type=["xlsx","xlsb"])
if not file:
    st.stop()

# detect file
if file.name.endswith(".xlsb"):
    excel_file = pd.ExcelFile(file, engine="pyxlsb")
else:
    excel_file = pd.ExcelFile(file)

sheet = st.selectbox("Pilih Sheet", excel_file.sheet_names)

if file.name.endswith(".xlsb"):
    df = pd.read_excel(file, sheet_name=sheet, engine="pyxlsb")
else:
    df = pd.read_excel(file, sheet_name=sheet)

# ===============================
# NORMALISASI
# ===============================
df.columns = df.columns.str.strip().str.lower()

mapping = {
    'prod':'PROD','cob':'COB','uy':'UY','curr':'CURRENCY',
    'broker':'BROKER',
    'qs_ceding':'QS_CEDING','sp_ceding':'SP_CEDING',
    'komisi_qs':'KOMISI_QS','komisi_sp':'KOMISI_SP',
    'klaim_qs':'KLAIM_QS','klaim_sp':'KLAIM_SP'
}
df = df.rename(columns=mapping)

for col in ['CURRENCY', 'COB', 'BROKER']:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.upper()

df = df[(df['CURRENCY'] != "") & (df['CURRENCY'].notna())]

# ===============================
# REF START
# ===============================
start_number = st.number_input("Start Ref No", value=81, step=1)
file_name = st.text_input("Nama file", "SOA_Report")
note = st.text_area("Note")

zero_option = st.selectbox("Tampilkan Baris Nol", ["Show All", "Hide Zero Rows"])

# ===============================
# CLEAN NUMBER
# ===============================
def clean_number(x):
    try:
        if pd.isna(x):
            return 0
        return float(str(x).replace(",", ""))
    except:
        return 0

num_cols = ['QS_CEDING','SP_CEDING','KOMISI_QS','KOMISI_SP','KLAIM_QS','KLAIM_SP']

for c in num_cols:
    if c in df.columns:
        df[c] = df[c].apply(clean_number)

# ===============================
# GENERATE REPORT
# ===============================
def generate_report(df, tipe, zero_option):

    if tipe == "QS":
        df['PREMIUM'] = df['QS_CEDING']
        df['COMMISSION'] = df['KOMISI_QS']
        df['CLAIM'] = df['KLAIM_QS']
    else:
        df['PREMIUM'] = df['SP_CEDING']
        df['COMMISSION'] = df['KOMISI_SP']
        df['CLAIM'] = df['KLAIM_SP']

    grouped = df.groupby(['CURRENCY','COB','UY']).sum(numeric_only=True).reset_index()

    grouped['AMOUNT'] = grouped['PREMIUM'] - grouped['COMMISSION'] - grouped['CLAIM']

    rows = []

    for curr, d1 in grouped.groupby('CURRENCY'):

        if zero_option == "Hide Zero Rows":
            d1 = d1[~(d1[['PREMIUM','COMMISSION','CLAIM','AMOUNT']] == 0).all(axis=1)]

        for cob, d2 in d1.groupby('COB'):

            first = True
            for _, r in d2.iterrows():

                rows.append([
                    curr if first else "",
                    cob if first else "",
                    r['UY'],
                    r['PREMIUM'],
                    r['COMMISSION'],
                    r['CLAIM'],
                    r['AMOUNT']
                ])
                first = False

            subtotal = d2[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
            rows.append(["", f"{cob} TOTAL", "", *subtotal])

        total = d1[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
        rows.append([f"{curr} TOTAL","","", *total])
        rows.append(["","","","","","",""])

    return pd.DataFrame(rows, columns=['CURRENCY','COB','UY','PREMIUM','COMMISSION','CLAIM','AMOUNT'])

# ===============================
# EXPORT WORD
# ===============================
def export_word(df, brokers, start_number, note, zero_option):

    doc = Document()
    ref_counter = start_number

    for broker in brokers:

        df_b = df[df["BROKER"] == broker]

        qs = generate_report(df_b.copy(), "QS", zero_option)
        sp = generate_report(df_b.copy(), "SP", zero_option)

        # QS PAGE
        doc.add_paragraph("STATEMENT OF ACCOUNT").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Ref No: {ref_counter}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"QUOTA SHARE - {broker}")

        ref_counter += 1

        table = doc.add_table(rows=1, cols=len(qs.columns))
        for i,c in enumerate(qs.columns):
            table.rows[0].cells[i].text = c

        for _, r in qs.iterrows():
            row = table.add_row().cells
            for i,v in enumerate(r):
                row[i].text = str(v)

        doc.add_page_break()

        # SP PAGE
        doc.add_paragraph("STATEMENT OF ACCOUNT").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Ref No: {ref_counter}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"SURPLUS - {broker}")

        ref_counter += 1

        table = doc.add_table(rows=1, cols=len(sp.columns))
        for i,c in enumerate(sp.columns):
            table.rows[0].cells[i].text = c

        for _, r in sp.iterrows():
            row = table.add_row().cells
            for i,v in enumerate(r):
                row[i].text = str(v)

        doc.add_page_break()

    doc.add_paragraph(f"Note: {note}")

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# ===============================
# BUTTONS
# ===============================
col1, col2 = st.columns(2)

with col1:
    if st.button("⬇️ Download Excel"):

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:

            brokers = df["BROKER"].dropna().unique()

            ref = start_number

            for b in brokers:
                df_b = df[df["BROKER"] == b]

                qs = generate_report(df_b.copy(), "QS", zero_option)
                sp = generate_report(df_b.copy(), "SP", zero_option)

                qs.to_excel(writer, sheet_name=f"QS_{b}"[:31], index=False)
                sp.to_excel(writer, sheet_name=f"SP_{b}"[:31], index=False)

        st.download_button(
            "Download Excel File",
            output.getvalue(),
            file_name=f"{file_name}.xlsx"
        )

with col2:
    if st.button("📄 Download Word"):

        brokers = df["BROKER"].dropna().unique()

        file_stream = export_word(
            df,
            brokers,
            start_number,
            note,
            zero_option
        )

        st.download_button(
            "Download Word File",
            file_stream,
            file_name=f"{file_name}.docx"
        )
