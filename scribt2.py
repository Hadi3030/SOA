import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="SOA Report", layout="wide")

st.title("📑 SOA Report Generator")

file = st.file_uploader("Upload File SOA", type=["xlsx"])

if not file:
    st.stop()

excel_file = pd.ExcelFile(file)
sheet_names = excel_file.sheet_names

# PILIH SHEET
sheet = st.selectbox("Pilih Sheet", sheet_names)

df = pd.read_excel(excel_file, sheet_name=sheet)

# ===============================
# CLEAN NUMERIC (PENTING BANGET)
# ===============================
num_cols = ['QS_CEDING','SP_CEDING','KOMISI_QS','KOMISI_SP']

for col in num_cols:
    df[col] = (
        df[col]
        .astype(str)
        .str.replace(',', '', regex=False)   # hapus koma ribuan
        .str.replace(' ', '', regex=False)   # hapus spasi
    )
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
df.columns = df.columns.str.strip().str.upper()

required_cols = [
    'PROD','BRANCH NAME','COB','POLICY NO.','UY','CURRENCY',
    'KURS','BROKER','SHARERE','QS_CEDING','KOMISI_QS',
    'SP_CEDING','KOMISI_SP'
]

missing = [col for col in required_cols if col not in df.columns]

if missing:
    st.error(f"Kolom tidak ditemukan: {missing}")
    st.stop()

# ===============================
# CALCULATION
# ===============================
df['PREMIUM'] = df['QS_CEDING'] + df['SP_CEDING']
df['COMMISSION'] = df['KOMISI_QS'] + df['KOMISI_SP']
df['CLAIM'] = 0
df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION']

# ===============================
# GROUPING
# ===============================
grouped = df.groupby(['CURRENCY','COB','UY']).agg({
    'PREMIUM':'sum',
    'COMMISSION':'sum',
    'CLAIM':'sum',
    'AMOUNT':'sum'
}).reset_index()

# ===============================
# FORMAT REPORT (WITH TOTAL)
# ===============================
final_rows = []

for curr in grouped['CURRENCY'].unique():
    df_curr = grouped[grouped['CURRENCY'] == curr]

    for cob in df_curr['COB'].unique():
        df_cob = df_curr[df_curr['COB'] == cob]

        for _, row in df_cob.iterrows():
            final_rows.append([
                curr, cob, row['UY'],
                row['PREMIUM'], row['COMMISSION'],
                row['CLAIM'], row['AMOUNT']
            ])

        total_cob = df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
        final_rows.append(["", f"{cob} TOTAL", "", *total_cob])

    total_curr = df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
    final_rows.append([f"{curr} TOTAL","","", *total_curr])

report = pd.DataFrame(final_rows,
    columns=['CURRENCY','COB','UW YEAR','PREMIUM','COMMISSION','CLAIM','AMOUNT'])

# ===============================
# DISPLAY
# ===============================
st.subheader("📊 Report SOA")
st.dataframe(report)

ref_no = st.text_input("Ref No", value="")
treaty_year = st.text_input("Treaty Year", value="2026")
quarter = st.text_input("Quarter", value="Q1")
months = st.text_input("For Months", value="Jan - Mar")
remarks = st.text_input("Remarks", value="-")

# ===============================
# EXPORT EXCEL
# ===============================
file_name = st.text_input("Nama file", value="SOA_Report")

output = io.BytesIO()

from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

with pd.ExcelWriter(output, engine='openpyxl') as writer:

    report.to_excel(writer, index=False, sheet_name='SOA Report', startrow=8)

    ws = writer.sheets['SOA Report']

    # HEADER
    ws.merge_cells('A2:G2')
    ws['A2'] = "STATEMENT OF ACCOUNT"
    ws['A2'].font = Font(bold=True, size=14)
    ws['A2'].alignment = Alignment(horizontal='center')

    # FORMAT NUMBER
    number_format = '#,##0.00;[Red](#,##0.00)'
    for col in ['D','E','F','G']:
        for row in range(10, ws.max_row + 1):
            ws[f"{col}{row}"].number_format = number_format

    # AUTO WIDTH
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)

        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_length + 2

st.download_button(
    "⬇️ Download Report",
    data=output.getvalue(),
    file_name=f"{file_name}.xlsx"
)
