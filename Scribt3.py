import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="SOA Report ALL Broker", layout="wide")
st.title("📑 SOA Report Generator - All Broker")

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
    'broker':'BROKER',
    'qs_ceding':'QS_CEDING','sp_ceding':'SP_CEDING',
    'komisi_qs':'KOMISI_QS','komisi_sp':'KOMISI_SP',
    'klaim_qs':'KLAIM_QS','klaim_sp':'KLAIM_SP'
}
df = df.rename(columns=mapping)

# ===============================
# FIX TEXT
# ===============================
for col in ['CURRENCY', 'COB', 'BROKER']:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.upper()

df = df[(df['CURRENCY'] != "") & (df['CURRENCY'].notna())]

# ===============================
# FILTER COB
# ===============================
st.subheader("Pilih COB")

if "COB" in df.columns:
    cob_list = sorted(df["COB"].dropna().unique().tolist())

    selected_cob = cob_list if st.checkbox("ALL COB", True) else [
        cob for cob in cob_list if st.checkbox(cob)
    ]

    df = df[df["COB"].isin(selected_cob)]

# ===============================
# FILTER UW YEAR
# ===============================
st.subheader("Pilih UW Year")

if "UY" in df.columns:
    uy_list = sorted(df["UY"].dropna().unique().tolist())

    selected_uy = uy_list if st.checkbox("ALL UW YEAR", True) else [
        uy for uy in uy_list if st.checkbox(str(uy))
    ]

    df = df[df["UY"].isin(selected_uy)]

# ===============================
# FILTER LT
# ===============================
lt_option = st.selectbox("Filter Long Term", ["ALL", "LT", "NON-LT"])

if lt_option != "ALL":
    if lt_option == "LT":
        df = df[df['COB'].str.contains("LT", na=False)]
    else:
        df = df[~df['COB'].str.contains("LT", na=False)]

zero_option = st.selectbox(
    "Tampilkan Baris Nol",
    ["Show All", "Hide Zero Rows"]
)

# ===============================
# CLEAN NUMERIC
# ===============================
num_cols = [
    'QS_CEDING','SP_CEDING',
    'KOMISI_QS','KOMISI_SP',
    'KLAIM_QS','KLAIM_SP'
]

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
# DATE INFO
# ===============================
month_map = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
             7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

min_m = int(df['MONTH'].min())
max_m = int(df['MONTH'].max())
year = int(df['YEAR'].mode()[0])

months_text = f"{month_map[min_m]} - {month_map[max_m]} {year}"

def get_quarter(m):
    return ["I","II","III","IV"][(m-1)//3]

quarter = get_quarter(max_m)

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

    df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION']

    grouped = (
        df.groupby(['CURRENCY','COB','UY'])
        .sum(numeric_only=True)
        .reset_index()
        .sort_values(['CURRENCY','COB','UY'])
    )

    rows = []

    for curr, df_curr in grouped.groupby('CURRENCY'):

        if zero_option == "Hide Zero Rows":
            df_curr = df_curr[
                ~((df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']] == 0).all(axis=1))
            ]

        if df_curr.empty:
            continue

        first_row = True

        for cob, df_cob in df_curr.groupby('COB'):

            if df_cob.empty:
                continue

            first_cob_row = True

            for _, r in df_cob.iterrows():
                rows.append([
                    curr if first_row else "",
                    cob if first_cob_row else "",
                    r['UY'],
                    r['PREMIUM'],
                    r['COMMISSION'],
                    r['CLAIM'],
                    r['AMOUNT']
                ])
                first_row = False
                first_cob_row = False

            subtotal = df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
            rows.append(["", f"{cob} TOTAL", "", *subtotal])

        total_curr = df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
        rows.append([f"{curr} TOTAL","","", *total_curr])
        rows.append(["","","","","","",""])

    return pd.DataFrame(
        rows,
        columns=['CURRENCY','COB','UW YEAR','PREMIUM','COMMISSION','CLAIM','AMOUNT']
    )

# ===============================
# INPUT USER
# ===============================
ref_qs = st.text_input("Ref No QS")
ref_sp = st.text_input("Ref No SL")
note = st.text_area("Note")
file_name = st.text_input("Nama file", value="SOA_Report")

# ===============================
# EXPORT MULTI BROKER
# ===============================
if st.button("⬇️ Download All Broker"):

    output = io.BytesIO()

    from openpyxl.styles import Font, Alignment, PatternFill, Border
    from openpyxl.drawing.image import Image

    header_fill = PatternFill("solid", fgColor="000000")
    grey_fill = PatternFill("solid", fgColor="D9D9D9")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    no_border = Border()

    def write_sheet(writer, data, name, tipe, ref, broker_name):

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
        ws['A10'] = "Broker       :"; ws['B10'] = broker_name

        # HEADER TABLE
        for col in "ABCDEFG":
            cell = ws[f"{col}13"]
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True)

        current_currency = None

        for row in range(14, ws.max_row+1):

            val_curr = ws[f"A{row}"].value
            val_cob  = ws[f"B{row}"].value

            for col in "ABCDEFG":
                ws[f"{col}{row}"].fill = white_fill

            if all(ws[f"{col}{row}"].value in ["", None] for col in "ABCDEFG"):
                current_currency = None
                continue

            if val_curr not in ["", None] and "TOTAL" not in str(val_curr):
                current_currency = val_curr

            if current_currency:
                ws[f"A{row}"].fill = grey_fill

            if val_curr and "TOTAL" in str(val_curr):
                for col in "ABCDEFG":
                    ws[f"{col}{row}"].fill = grey_fill
                current_currency = None

            for col in ['D','E','F','G']:
                ws[f"{col}{row}"].number_format = '#,##0.00;[Red](#,##0.00)'

        # NOTE
        last = ws.max_row + 2
        ws[f"A{last}"] = "Note :"
        ws[f"B{last}"] = note

    with pd.ExcelWriter(output, engine='openpyxl') as writer:

        brokers = df["BROKER"].dropna().unique()

        for broker in brokers:
            df_b = df[df["BROKER"] == broker]

            report_qs = generate_report(df_b.copy(), "QS", zero_option)
            report_sp = generate_report(df_b.copy(), "SP", zero_option)

            write_sheet(writer, report_qs, f"QS_{broker}"[:31], "Quota Share", ref_qs, broker)
            write_sheet(writer, report_sp, f"SP_{broker}"[:31], "Surplus", ref_sp, broker)

    st.download_button(
        "📥 Download Excel",
        data=output.getvalue(),
        file_name=f"{file_name}_ALL.xlsx"
    )
