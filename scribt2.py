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

