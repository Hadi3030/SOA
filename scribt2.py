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
    'prod':'PROD','cob':'COB','uy':'UY','curr':'CURRENCY',
    'broker':'BROKER','qs_ceding':'QS_CEDING','sp_ceding':'SP_CEDING',
    'komisi_qs':'KOMISI_QS','komisi_sp':'KOMISI_SP'
}
df = df.rename(columns=mapping)

# ===============================
# FILTER BROKER
# ===============================
broker_list = df['BROKER'].dropna().unique().tolist()
broker_list.insert(0, "ALL")
selected_broker = st.selectbox("Pilih Broker", broker_list)

if selected_broker != "ALL":
    df = df[df['BROKER'] == selected_broker]

# ===============================
# FILTER LT
# ===============================
lt_option = st.selectbox("Filter Long Term", ["ALL", "LT", "NON-LT"])

if lt_option != "ALL":
    if lt_option == "LT":
        df = df[df['COB'].str.contains("LT", case=False, na=False)]
    else:
        df = df[~df['COB'].str.contains("LT", case=False, na=False)]

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

def get_quarter(m):
    if m <= 3: return "I"
    elif m <= 6: return "II"
    elif m <= 9: return "III"
    else: return "IV"

quarter = get_quarter(max_m)

# ===============================
# GENERATE REPORT
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

    rows = []

    for curr in grouped['CURRENCY'].unique():
        df_curr = grouped[grouped['CURRENCY'] == curr]

        first_row = True

        for cob in df_curr['COB'].unique():
            df_cob = df_curr[df_curr['COB'] == cob]

            for _, r in df_cob.iterrows():
                rows.append([
                    curr if first_row else "",
                    cob,
                    r['UY'],
                    r['PREMIUM'],
                    r['COMMISSION'],
                    r['CLAIM'],
                    r['AMOUNT']
                ])
                first_row = False

            subtotal = df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
            rows.append(["", f"{cob} TOTAL", "", *subtotal])

        total_curr = df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
        rows.append([f"{curr} TOTAL","","", *total_curr])
        rows.append(["","","","","","",""])  # spasi

    return pd.DataFrame(rows,
        columns=['CURRENCY','COB','UW YEAR','PREMIUM','COMMISSION','CLAIM','AMOUNT'])

report_qs = generate_report(df.copy(), "QS")
report_sp = generate_report(df.copy(), "SP")

st.dataframe(report_qs)

# ===============================
# INPUT
# ===============================
ref_qs = st.text_input("Ref No QS")
ref_sp = st.text_input("Ref No SL")
note = st.text_area("Note")
file_name = st.text_input("Nama file", value="SOA_Report")

# ===============================
# EXPORT
# ===============================
output = io.BytesIO()

from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.drawing.image import Image

header_fill = PatternFill("solid", fgColor="000000")
grey_fill = PatternFill("solid", fgColor="D9D9D9")

def write_sheet(writer, data, name, tipe, ref):

    data.to_excel(writer, index=False, sheet_name=name, startrow=12)
    ws = writer.sheets[name]

    # LOGO
    try:
        logo = Image("askrindo.jpg")
        logo.height = 60
        logo.width = 140
        ws.add_image(logo, "A1")
    except:
        pass

    # TITLE
    ws.merge_cells('A4:G4')
    ws['A4'] = "STATEMENT OF ACCOUNT"
    ws['A4'].font = Font(bold=True, size=14)
    ws['A4'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A5:G5')
    ws['A5'] = f"Ref No. {ref}"
    ws['A5'].font = Font(bold=True)
    ws['A5'].alignment = Alignment(horizontal='center')

    # HEADER INFO
    ws['A7'] = "Treaty Year  :"; ws['B7'] = year
    ws['A8'] = "Quarter      :"; ws['B8'] = f"{quarter} {tipe}"
    ws['A9'] = "For Months   :"; ws['B9'] = months_text
    ws['A10'] = "Broker       :"; ws['B10'] = selected_broker

    # STYLE HEADER TABLE
    for col in "ABCDEFG":
        cell = ws[f"{col}13"]
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)

    # FORMAT BODY
    for row in range(14, ws.max_row+1):

        ws[f"A{row}"].alignment = Alignment(horizontal='left')
        ws[f"B{row}"].alignment = Alignment(horizontal='left')
        ws[f"C{row}"].alignment = Alignment(horizontal='center')

        # Currency shading
        if ws[f"A{row}"].value not in ["", None]:
            for col in "ABCDEFG":
                ws[f"{col}{row}"].fill = grey_fill

        # TOTAL bold
        if "TOTAL" in str(ws[f"B{row}"].value) or "TOTAL" in str(ws[f"A{row}"].value):
            for col in "ABCDEFG":
                ws[f"{col}{row}"].font = Font(bold=True)

        # NUMBER FORMAT
        for col in ['D','E','F','G']:
            ws[f"{col}{row}"].number_format = '#,##0.00;[Red](#,##0.00)'

    # NOTE
    last = ws.max_row + 2
    ws[f"A{last}"] = "Note :"
    ws[f"A{last}"].font = Font(bold=True)

    ws[f"B{last}"] = note
    ws[f"B{last}"].font = Font(bold=True)

with pd.ExcelWriter(output, engine='openpyxl') as writer:
    write_sheet(writer, report_qs, "QS Report", "Quota Share", ref_qs)
    write_sheet(writer, report_sp, "SL Report", "Surplus", ref_sp)

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
# # UPLOAD
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
# broker_list = df['BROKER'].dropna().unique().tolist()
# broker_list.insert(0, "ALL")

# selected_broker = st.selectbox("Pilih Broker", broker_list)

# if selected_broker != "ALL":
#     df = df[df['BROKER'] == selected_broker]

# # ===============================
# # SELECT LT
# # ===============================
# lt_option = st.selectbox("Filter Long Term", ["ALL", "LT", "NON-LT"])

# if lt_option != "ALL":
#     if lt_option == "LT":
#         df = df[df['COB'].str.contains("LT", case=False, na=False)]
#     else:
#         df = df[~df['COB'].str.contains("LT", case=False, na=False)]

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
# # PARSE PROD
# # ===============================
# def parse_prod(x):
#     try:
#         x = str(x)
#         return int(x[:4]), int(x[-2:])
#     except:
#         return None, None

# df[['YEAR','MONTH']] = df['PROD'].apply(lambda x: pd.Series(parse_prod(x)))
# df = df.dropna(subset=['YEAR','MONTH'])

# # ===============================
# # MONTH + QUARTER
# # ===============================
# month_map = {
#     1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
#     7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'
# }

# min_m = int(df['MONTH'].min())
# max_m = int(df['MONTH'].max())
# year = int(df['YEAR'].mode()[0])

# months_text = f"{month_map[min_m]} - {month_map[max_m]} {year}"

# def get_quarter(m):
#     if m <= 3: return "I"
#     elif m <= 6: return "II"
#     elif m <= 9: return "III"
#     else: return "IV"

# quarter = get_quarter(max_m)

# # ===============================
# # GENERATE REPORT + TOTAL
# # ===============================
# def generate_report(df, tipe):

#     if tipe == "QS":
#         df['PREMIUM'] = df['QS_CEDING']
#         df['COMMISSION'] = df['KOMISI_QS']
#     else:
#         df['PREMIUM'] = df['SP_CEDING']
#         df['COMMISSION'] = df['KOMISI_SP']

#     df['CLAIM'] = 0
#     df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION']

#     grouped = df.groupby(['CURRENCY','COB','UY']).sum(numeric_only=True).reset_index()

#     rows = []

#     for curr in grouped['CURRENCY'].unique():
#         df_curr = grouped[grouped['CURRENCY'] == curr]

#         first_currency_row = True

#         for cob in df_curr['COB'].unique():
#             df_cob = df_curr[df_curr['COB'] == cob]

#             for _, r in df_cob.iterrows():

#                 currency_val = curr if first_currency_row else ""

#                 rows.append([
#                     currency_val,
#                     cob,
#                     r['UY'],
#                     r['PREMIUM'],
#                     r['COMMISSION'],
#                     r['CLAIM'],
#                     r['AMOUNT']
#                 ])

#                 first_currency_row = False

#             # SUBTOTAL PER COB
#             subtotal = df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
#             rows.append(["", f"{cob} TOTAL", "", *subtotal])

#         # TOTAL PER CURRENCY
#         total_curr = df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
#         rows.append([f"{curr} TOTAL","","", *total_curr])

#         # SPASI ANTAR CURRENCY
#         rows.append(["","","","","","",""])

#     return pd.DataFrame(
#         rows,
#         columns=['CURRENCY','COB','UW YEAR','PREMIUM','COMMISSION','CLAIM','AMOUNT']
#     )
    
# report_qs = generate_report(df.copy(), "QS")
# report_sp = generate_report(df.copy(), "SP")

# st.dataframe(report_qs)

# # ===============================
# # INPUT
# # ===============================
# ref_qs = st.text_input("Ref No QS")
# ref_sp = st.text_input("Ref No SL")
# note = st.text_area("Note")
# file_name = st.text_input("Nama file", value="SOA_Report")

# # ===============================
# # EXPORT
# # ===============================
# output = io.BytesIO()

# from openpyxl.styles import Font, Alignment
# from openpyxl.drawing.image import Image

# def write_sheet(writer, data, name, tipe, ref):

#     data.to_excel(writer, index=False, sheet_name=name, startrow=12)
#     ws = writer.sheets[name]

#     # LOGO (fix 3 baris)
#     try:
#         logo = Image("askrindo.jpg")
#         logo.height = 60
#         logo.width = 140
#         ws.add_image(logo, "A1")
#     except:
#         pass

#     # TITLE (baris 4)
#     ws.merge_cells('A4:G4')
#     ws['A4'] = "STATEMENT OF ACCOUNT"
#     ws['A4'].font = Font(bold=True, size=14)
#     ws['A4'].alignment = Alignment(horizontal='center')

#     ws.merge_cells('A5:G5')
#     ws['A5'] = f"Ref No. {ref}"
#     ws['A5'].font = Font(bold=True)
#     ws['A5'].alignment = Alignment(horizontal='center')

#     # HEADER RAPI
#     ws['A7'] = "Treaty Year  :"
#     ws['B7'] = year

#     ws['A8'] = "Quarter      :"
#     ws['B8'] = f"{quarter} {tipe}"

#     ws['A9'] = "For Months   :"
#     ws['B9'] = months_text

#     ws['A10'] = "Broker       :"
#     ws['B10'] = selected_broker

#     # BOLD COLUMN
#     for row in range(13, ws.max_row+1):
#         ws[f"A{row}"].font = Font(bold=True)
#         ws[f"B{row}"].font = Font(bold=True)

#         if "TOTAL" in str(ws[f"B{row}"].value) or "TOTAL" in str(ws[f"A{row}"].value):
#             for col in "ABCDEFG":
#                 ws[f"{col}{row}"].font = Font(bold=True)

#     # NOTE PALING BAWAH
#     last_row = ws.max_row + 2
#     ws[f"A{last_row}"] = "Note :"
#     ws[f"A{last_row}"].font = Font(bold=True)

#     ws[f"B{last_row}"] = note
#     ws[f"B{last_row}"].font = Font(bold=True)

# with pd.ExcelWriter(output, engine='openpyxl') as writer:
#     write_sheet(writer, report_qs, "QS Report", "Quota Share", ref_qs)
#     write_sheet(writer, report_sp, "SL Report", "Surplus", ref_sp)

# st.download_button(
#     "⬇️ Download Report",
#     data=output.getvalue(),
#     file_name=f"{file_name}.xlsx"
# )
