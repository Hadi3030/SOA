import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="SOA Report", layout="wide")
st.title("📑 SOA Report Generator")

# ===============================
# UPLOAD
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
    'prod':'PROD','cob':'COB','uy':'UY',
    'curr':'CURRENCY','broker':'BROKER',
    'qs_ceding':'QS_CEDING','sp_ceding':'SP_CEDING',
    'komisi_qs':'KOMISI_QS','komisi_sp':'KOMISI_SP'
}
df = df.rename(columns=mapping)

# ===============================
# SELECT BROKER
# ===============================
broker_list = sorted(df['BROKER'].dropna().unique())
selected_broker = st.multiselect(
    "Pilih Broker",
    options=["ALL"] + broker_list,
    default=["ALL"]
)

if "ALL" not in selected_broker:
    df = df[df['BROKER'].isin(selected_broker)]

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
# PARSE PROD
# ===============================
def parse_prod(x):
    try:
        x = str(x)
        return int(x[:4]), int(x[-2:])
    except:
        return None, None

df[['YEAR','MONTH']] = df['PROD'].apply(lambda x: pd.Series(parse_prod(x)))
df = df.dropna(subset=['YEAR','MONTH'])

# ===============================
# MONTH + QUARTER
# ===============================
month_map = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
             7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

min_m = int(df['MONTH'].min())
max_m = int(df['MONTH'].max())
year = int(df['YEAR'].mode()[0])

months_text = f"{month_map[min_m]} - {month_map[max_m]} {year}"

def get_q(m):
    return ["I","II","III","IV"][(m-1)//3]

quarter = get_q(max_m)

# ===============================
# GENERATE REPORT (WITH TOTAL)
# ===============================
def generate_report(df, tipe):

    if tipe == "QS":
        df['PREMIUM'] = df['QS_CEDING']
        df['COMMISSION'] = df['KOMISI_QS']
    else:
        df['PREMIUM'] = df['SP_CEDING']
        df['COMMISSION'] = df['KOMISI_SP']

    df['CLAIM'] = 0
    df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION']

    grouped = df.groupby(['CURRENCY','COB','UY']).sum(numeric_only=True).reset_index()

    final = []

    for curr in grouped['CURRENCY'].unique():
        df_curr = grouped[grouped['CURRENCY']==curr]

        for cob in df_curr['COB'].unique():
            df_cob = df_curr[df_curr['COB']==cob]

            for _, r in df_cob.iterrows():
                final.append([curr, cob, r['UY'], r['PREMIUM'], r['COMMISSION'], r['CLAIM'], r['AMOUNT']])

            # TOTAL COB
            t = df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
            final.append(["", f"{cob} TOTAL", "", *t])

        # TOTAL CURRENCY
        t2 = df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
        final.append([f"{curr} TOTAL","","",*t2])

        # SPASI
        final.append(["","","","","","",""])

    return pd.DataFrame(final, columns=[
        'CURRENCY','COB','UW YEAR','PREMIUM','COMMISSION','CLAIM','AMOUNT'
    ])

report_qs = generate_report(df.copy(),"QS")
report_sl = generate_report(df.copy(),"SL")

st.subheader("QS Report")
st.dataframe(report_qs)

st.subheader("SL Report")
st.dataframe(report_sl)

# ===============================
# INPUT
# ===============================
ref_qs = st.text_input("Ref No QS")
ref_sl = st.text_input("Ref No SL")
note = st.text_area("Note")
file_name = st.text_input("Nama file","SOA_Report")

# ===============================
# EXPORT
# ===============================
output = io.BytesIO()

from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.drawing.image import Image

def write_sheet(writer, data, sheet, tipe, ref):

    data.to_excel(writer, index=False, sheet_name=sheet, startrow=12)
    ws = writer.sheets[sheet]

    # LOGO
    try:
        logo = Image("askrindo.jpg")
        logo.height = 80
        ws.add_image(logo, "A1")
    except:
        pass

    # TITLE (ROW 4)
    ws.merge_cells('A4:G4')
    ws['A4'] = "STATEMENT OF ACCOUNT"
    ws['A4'].font = Font(bold=True, size=14)
    ws['A4'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A5:G5')
    ws['A5'] = f"Ref No. {ref}"
    ws['A5'].alignment = Alignment(horizontal='center')

    # INFO RAPI (SATU KOLOM TITIK DUA)
    ws['A7'] = "Treaty Year  :"
    ws['B7'] = year

    ws['A8'] = "Quarter      :"
    ws['B8'] = f"{quarter} {tipe}"

    ws['A9'] = "For Months   :"
    ws['B9'] = months_text

    ws['A10'] = "Broker       :"
    ws['B10'] = "ALL" if "ALL" in selected_broker else ", ".join(selected_broker)

    ws['A11'] = "Note         :"
    ws['B11'] = note

    # HEADER STYLE
    fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)

    for col in range(1,8):
        c = ws.cell(row=13,column=col)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal='center')

    # FORMAT ANGKA
    for col in ['D','E','F','G']:
        for r in range(14, ws.max_row+1):
            ws[f"{col}{r}"].number_format = '#,##0.00;[Red](#,##0.00)'

# WRITE
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    write_sheet(writer, report_qs, "QS Report", "Quota Share", ref_qs)
    write_sheet(writer, report_sl, "SL Report", "Surplus", ref_sl)

# DOWNLOAD
st.download_button(
    "⬇️ Download",
    data=output.getvalue(),
    file_name=f"{file_name}.xlsx"
)

# import streamlit as st
# import pandas as pd
# import io

# st.set_page_config(page_title="SOA Report", layout="wide")
# st.title("📑 SOA Report Generator")

# # ===============================
# # UPLOAD FILE
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

# mapping = {
#     'prod':'PROD',
#     'cob':'COB',
#     'uy':'UY',
#     'curr':'CURRENCY',
#     'broker':'BROKER',
#     'qs_ceding':'QS_CEDING',
#     'sp_ceding':'SP_CEDING',
#     'komisi_qs':'KOMISI_QS',
#     'komisi_sp':'KOMISI_SP'
# }
# df = df.rename(columns=mapping)

# # ===============================
# # SELECT BROKER
# # ===============================
# broker_list = df['BROKER'].dropna().unique()
# selected_broker = st.selectbox("Pilih Broker", broker_list)

# df = df[df['BROKER'] == selected_broker]

# # ===============================
# # CLEAN NUMERIC
# # ===============================
# num_cols = ['QS_CEDING','SP_CEDING','KOMISI_QS','KOMISI_SP']

# for col in num_cols:
#     df[col] = (
#         df[col].astype(str)
#         .str.replace('.', '', regex=False)
#         .str.replace(',', '.', regex=False)
#     )
#     df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# # ===============================
# # PARSE PROD → BULAN
# # ===============================
# def parse_prod(x):
#     try:
#         x = str(x).strip()
#         year = int(x[:4])
#         month = int(x[-2:])
#         return year, month
#     except:
#         return None, None

# df[['YEAR','MONTH']] = df['PROD'].apply(lambda x: pd.Series(parse_prod(x)))

# # DROP NULL
# df = df.dropna(subset=['YEAR','MONTH'])

# # ===============================
# # MONTH RANGE
# # ===============================
# month_map = {
#     1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
#     7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'
# }

# min_month = int(df['MONTH'].min())
# max_month = int(df['MONTH'].max())
# year = int(df['YEAR'].mode()[0])

# months_text = f"{month_map[min_month]} - {month_map[max_month]} {year}"

# # ===============================
# # QUARTER DETECTION
# # ===============================
# def get_quarter(m):
#     if m <= 3: return "I"
#     elif m <= 6: return "II"
#     elif m <= 9: return "III"
#     else: return "IV"

# quarter = get_quarter(max_month)

# # ===============================
# # CALCULATION
# # ===============================
# df['PREMIUM_QS'] = df['QS_CEDING']
# df['PREMIUM_SP'] = df['SP_CEDING']

# df['COMMISSION_QS'] = df['KOMISI_QS']
# df['COMMISSION_SP'] = df['KOMISI_SP']

# # ===============================
# # FUNCTION REPORT
# # ===============================
# def generate_report(df, tipe):

#     if tipe == "QS":
#         df['PREMIUM'] = df['PREMIUM_QS']
#         df['COMMISSION'] = df['COMMISSION_QS']
#     else:
#         df['PREMIUM'] = df['PREMIUM_SP']
#         df['COMMISSION'] = df['COMMISSION_SP']

#     df['CLAIM'] = 0
#     df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION']

#     grouped = df.groupby(['CURRENCY','COB','UY']).agg({
#         'PREMIUM':'sum',
#         'COMMISSION':'sum',
#         'CLAIM':'sum',
#         'AMOUNT':'sum'
#     }).reset_index()

#     return grouped

# report_qs = generate_report(df.copy(), "QS")
# report_sp = generate_report(df.copy(), "SP")

# st.subheader("📊 QS Report")
# st.dataframe(report_qs)

# st.subheader("📊 SL Report")
# st.dataframe(report_sp)

# # ===============================
# # INPUT
# # ===============================
# ref_no = st.text_input("Ref No", value="")
# file_name = st.text_input("Nama file", value="SOA_Report")

# # ===============================
# # EXPORT
# # ===============================
# output = io.BytesIO()

# from openpyxl.styles import Font, Alignment
# from openpyxl.drawing.image import Image

# def write_sheet(writer, data, sheet_name, tipe):

#     data.to_excel(writer, index=False, sheet_name=sheet_name, startrow=10)
#     ws = writer.sheets[sheet_name]

#     try:
#         logo = Image("askrindo.jpg")
#         ws.add_image(logo, "A1")
#     except:
#         pass

#     ws.merge_cells('A2:G2')
#     ws['A2'] = "STATEMENT OF ACCOUNT"
#     ws['A2'].font = Font(bold=True, size=14)
#     ws['A2'].alignment = Alignment(horizontal='center')

#     ws.merge_cells('A3:G3')
#     ws['A3'] = f"Ref No. {ref_no}"
#     ws['A3'].alignment = Alignment(horizontal='center')

#     ws['A5'] = f"Treaty Year : {year}"
#     ws['A6'] = f"Quarter     : {quarter} {tipe}"
#     ws['A7'] = f"For Months  : {months_text}"
#     ws['A8'] = f"Broker      : {selected_broker}"

# with pd.ExcelWriter(output, engine='openpyxl') as writer:
#     write_sheet(writer, report_qs, "QS Report", "Quota Share")
#     write_sheet(writer, report_sp, "SL Report", "Surplus")

# st.download_button(
#     "⬇️ Download Report",
#     data=output.getvalue(),
#     file_name=f"{file_name}.xlsx"
# )

