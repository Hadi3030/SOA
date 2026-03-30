import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="SOA Report", layout="wide")
st.title("📑 SOA Report Generator")

# ===============================
# UPLOAD FILE
# ===============================
file = st.file_uploader("Upload File SOA", type=["xlsx"])
if not file:
    st.stop()

excel_file = pd.ExcelFile(file)
sheet = st.selectbox("Pilih Sheet", excel_file.sheet_names)
df = pd.read_excel(excel_file, sheet_name=sheet)

# ===============================
# NORMALISASI KOLOM
# ===============================
df.columns = df.columns.str.strip().str.lower()

mapping = {
    'prod':'PROD',
    'cob':'COB',
    'uy':'UY',
    'curr':'CURRENCY',
    'broker':'BROKER',
    'qs_ceding':'QS_CEDING',
    'sp_ceding':'SP_CEDING',
    'komisi_qs':'KOMISI_QS',
    'komisi_sp':'KOMISI_SP'
}
df = df.rename(columns=mapping)

# ===============================
# SELECT BROKER
# ===============================
broker_list = df['BROKER'].dropna().unique()
selected_broker = st.selectbox("Pilih Broker", broker_list)

df = df[df['BROKER'] == selected_broker]

# ===============================
# CLEAN NUMERIC
# ===============================
num_cols = ['QS_CEDING','SP_CEDING','KOMISI_QS','KOMISI_SP']

for col in num_cols:
    df[col] = (
        df[col].astype(str)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
    )
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# ===============================
# PARSE PROD → BULAN
# ===============================
def parse_prod(x):
    try:
        x = str(x).strip()
        year = int(x[:4])
        month = int(x[-2:])
        return year, month
    except:
        return None, None

df[['YEAR','MONTH']] = df['PROD'].apply(lambda x: pd.Series(parse_prod(x)))

# DROP NULL
df = df.dropna(subset=['YEAR','MONTH'])

# ===============================
# MONTH RANGE
# ===============================
month_map = {
    1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
    7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'
}

min_month = int(df['MONTH'].min())
max_month = int(df['MONTH'].max())
year = int(df['YEAR'].mode()[0])

months_text = f"{month_map[min_month]} - {month_map[max_month]} {year}"

# ===============================
# QUARTER DETECTION
# ===============================
def get_quarter(m):
    if m <= 3: return "I"
    elif m <= 6: return "II"
    elif m <= 9: return "III"
    else: return "IV"

quarter = get_quarter(max_month)

# ===============================
# CALCULATION
# ===============================
df['PREMIUM_QS'] = df['QS_CEDING']
df['PREMIUM_SP'] = df['SP_CEDING']

df['COMMISSION_QS'] = df['KOMISI_QS']
df['COMMISSION_SP'] = df['KOMISI_SP']

# ===============================
# FUNCTION REPORT
# ===============================
def generate_report(df, tipe):

    if tipe == "QS":
        df['PREMIUM'] = df['PREMIUM_QS']
        df['COMMISSION'] = df['COMMISSION_QS']
    else:
        df['PREMIUM'] = df['PREMIUM_SP']
        df['COMMISSION'] = df['COMMISSION_SP']

    df['CLAIM'] = 0
    df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION']

    grouped = df.groupby(['CURRENCY','COB','UY']).agg({
        'PREMIUM':'sum',
        'COMMISSION':'sum',
        'CLAIM':'sum',
        'AMOUNT':'sum'
    }).reset_index()

    return grouped

report_qs = generate_report(df.copy(), "QS")
report_sp = generate_report(df.copy(), "SP")

st.subheader("📊 QS Report")
st.dataframe(report_qs)

st.subheader("📊 SL Report")
st.dataframe(report_sp)

# ===============================
# INPUT
# ===============================
ref_no = st.text_input("Ref No", value="")
file_name = st.text_input("Nama file", value="SOA_Report")

# ===============================
# EXPORT
# ===============================
output = io.BytesIO()

from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image

def write_sheet(writer, data, sheet_name, tipe):

    data.to_excel(writer, index=False, sheet_name=sheet_name, startrow=10)
    ws = writer.sheets[sheet_name]

    try:
        logo = Image("askrindo.jpg")
        ws.add_image(logo, "A1")
    except:
        pass

    ws.merge_cells('A2:G2')
    ws['A2'] = "STATEMENT OF ACCOUNT"
    ws['A2'].font = Font(bold=True, size=14)
    ws['A2'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A3:G3')
    ws['A3'] = f"Ref No. {ref_no}"
    ws['A3'].alignment = Alignment(horizontal='center')

    ws['A5'] = f"Treaty Year : {year}"
    ws['A6'] = f"Quarter     : {quarter} {tipe}"
    ws['A7'] = f"For Months  : {months_text}"
    ws['A8'] = f"Broker      : {selected_broker}"

with pd.ExcelWriter(output, engine='openpyxl') as writer:
    write_sheet(writer, report_qs, "QS Report", "Quota Share")
    write_sheet(writer, report_sp, "SL Report", "Surplus")

st.download_button(
    "⬇️ Download Report",
    data=output.getvalue(),
    file_name=f"{file_name}.xlsx"
)



# import streamlit as st
# import pandas as pd
# import io

# st.set_page_config(page_title="SOA Report", layout="wide")
# st.title("📑 SOA Report Generator")

# # ===============================
# # UPLOAD FILE + PILIH SHEET
# # ===============================
# file = st.file_uploader("Upload File SOA", type=["xlsx"])

# if not file:
#     st.stop()

# excel_file = pd.ExcelFile(file)
# sheet = st.selectbox("Pilih Sheet", excel_file.sheet_names)
# df = pd.read_excel(excel_file, sheet_name=sheet)

# # ===============================
# # NORMALISASI KOLOM
# # ===============================
# df.columns = df.columns.str.strip().str.lower()

# column_mapping = {
#     'prod': 'PROD',
#     'cabang': 'BRANCH',
#     'cob': 'COB',
#     'policyno': 'POLICY_NO',
#     'policy no.': 'POLICY_NO',
#     'uy': 'UY',
#     'curr': 'CURRENCY',
#     'currency': 'CURRENCY',
#     'curr_us': 'KURS',
#     'rate': 'KURS',
#     'broker': 'BROKER',
#     'sharere': 'SHARERE',
#     'qs_ceding': 'QS_CEDING',
#     'komisi_qs': 'KOMISI_QS',
#     'sp_ceding': 'SP_CEDING',
#     'komisi_sp': 'KOMISI_SP',
#     'premi_reas': 'PREMI_REAS',
#     'komisi_reas': 'KOMISI_REAS',
#     'klaim_reas': 'KLAIM_REAS'
# }

# df = df.rename(columns=column_mapping)

# # ===============================
# # VALIDASI KOLOM WAJIB
# # ===============================
# required_cols = [
#     'COB','UY','CURRENCY',
#     'QS_CEDING','SP_CEDING',
#     'KOMISI_QS','KOMISI_SP'
# ]

# missing = [col for col in required_cols if col not in df.columns]

# if missing:
#     st.error(f"Kolom tidak ditemukan: {missing}")
#     st.write("Kolom terbaca:", list(df.columns))
#     st.stop()

# # ===============================
# # CLEAN NUMERIC
# # ===============================
# num_cols = ['QS_CEDING','SP_CEDING','KOMISI_QS','KOMISI_SP']

# for col in num_cols:
#     df[col] = (
#         df[col]
#         .astype(str)
#         .str.replace('.', '', regex=False)
#         .str.replace(',', '.', regex=False)
#         .str.replace(' ', '', regex=False)
#     )
#     df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# # ===============================
# # CALCULATION (SMART)
# # ===============================
# if 'PREMI_REAS' in df.columns:
#     df['PREMIUM'] = pd.to_numeric(df['PREMI_REAS'], errors='coerce').fillna(0)
# else:
#     df['PREMIUM'] = df['QS_CEDING'] + df['SP_CEDING']

# if 'KOMISI_REAS' in df.columns:
#     df['COMMISSION'] = pd.to_numeric(df['KOMISI_REAS'], errors='coerce').fillna(0)
# else:
#     df['COMMISSION'] = df['KOMISI_QS'] + df['KOMISI_SP']

# if 'KLAIM_REAS' in df.columns:
#     df['CLAIM'] = pd.to_numeric(df['KLAIM_REAS'], errors='coerce').fillna(0)
# else:
#     df['CLAIM'] = 0

# df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION'] - df['CLAIM']

# # ===============================
# # GROUPING
# # ===============================
# grouped = df.groupby(['CURRENCY','COB','UY']).agg({
#     'PREMIUM':'sum',
#     'COMMISSION':'sum',
#     'CLAIM':'sum',
#     'AMOUNT':'sum'
# }).reset_index()

# # ===============================
# # FORMAT REPORT + TOTAL
# # ===============================
# final_rows = []

# for curr in grouped['CURRENCY'].unique():
#     df_curr = grouped[grouped['CURRENCY'] == curr]

#     for cob in df_curr['COB'].unique():
#         df_cob = df_curr[df_curr['COB'] == cob]

#         for _, row in df_cob.iterrows():
#             final_rows.append([
#                 curr, cob, row['UY'],
#                 row['PREMIUM'], row['COMMISSION'],
#                 row['CLAIM'], row['AMOUNT']
#             ])

#         total_cob = df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
#         final_rows.append(["", f"{cob} TOTAL", "", *total_cob])

#     total_curr = df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
#     final_rows.append([f"{curr} TOTAL","","", *total_curr])

# report = pd.DataFrame(final_rows,
#     columns=['CURRENCY','COB','UW YEAR','PREMIUM','COMMISSION','CLAIM','AMOUNT'])

# # ===============================
# # DISPLAY
# # ===============================
# st.subheader("📊 Report SOA")
# st.dataframe(report)

# # ===============================
# # INPUT HEADER
# # ===============================
# st.subheader("📝 Informasi Laporan")

# ref_no = st.text_input("Ref No", value="")
# treaty_year = st.text_input("Treaty Year", value="2026")
# quarter = st.text_input("Quarter", value="Q1")
# months = st.text_input("For Months", value="Jan - Mar")
# remarks = st.text_input("Remarks", value="-")
# file_name = st.text_input("Nama file", value="SOA_Report")

# # ===============================
# # EXPORT EXCEL
# # ===============================
# output = io.BytesIO()

# from openpyxl.styles import Font, Alignment
# from openpyxl.drawing.image import Image
# from openpyxl.utils import get_column_letter

# with pd.ExcelWriter(output, engine='openpyxl') as writer:

#     report.to_excel(writer, index=False, sheet_name='SOA Report', startrow=10)

#     ws = writer.sheets['SOA Report']

#     # LOGO
#     try:
#         logo = Image("askrindo.jpg")
#         logo.width = 120
#         logo.height = 60
#         ws.add_image(logo, "A1")
#     except:
#         st.warning("Logo tidak ditemukan")

#     # HEADER
#     ws.merge_cells('A2:G2')
#     ws['A2'] = "STATEMENT OF ACCOUNT"
#     ws['A2'].font = Font(bold=True, size=14)
#     ws['A2'].alignment = Alignment(horizontal='center')

#     ws.merge_cells('A3:G3')
#     ws['A3'] = f"Ref No. {ref_no}"
#     ws['A3'].alignment = Alignment(horizontal='center')

#     ws['A5'] = f"Treaty Year : {treaty_year}"
#     ws['A6'] = f"Quarter     : {quarter}"
#     ws['A7'] = f"For Months  : {months}"
#     ws['A8'] = f"Remarks     : {remarks}"

#     # FORMAT ANGKA
#     number_format = '#,##0.00;[Red](#,##0.00)'
#     for col in ['D','E','F','G']:
#         for row in range(12, ws.max_row + 1):
#             ws[f"{col}{row}"].number_format = number_format

#     # AUTO WIDTH
#     for col in ws.columns:
#         max_length = 0
#         col_letter = get_column_letter(col[0].column)

#         for cell in col:
#             if cell.value:
#                 max_length = max(max_length, len(str(cell.value)))

#         ws.column_dimensions[col_letter].width = max_length + 2

# # ===============================
# # DOWNLOAD
# # ===============================
# st.download_button(
#     "⬇️ Download Report",
#     data=output.getvalue(),
#     file_name=f"{file_name}.xlsx"
# )


