import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="SOA Report ALL Broker", layout="wide")
st.title("📑 SOA Report Generator - All Broker")

# ===============================
# UPLOAD FILE
# ===============================
file = st.file_uploader("Upload File SOA", type=["xlsx"])

if not file:
    st.stop()

excel_file = pd.ExcelFile(file)
sheet = st.selectbox("Pilih Sheet", excel_file.sheet_names)

df = pd.read_excel(file, sheet_name=sheet)

# ===============================
# SIDEBAR FILTER
# ===============================
st.sidebar.header("Filter Data")

# ================= COB (CHECKBOX) =================
if "COB" in df.columns:
    cob_options = sorted(df["COB"].dropna().unique().tolist())
    selected_cob = st.sidebar.multiselect("Pilih COB", cob_options, default=cob_options)

    if selected_cob:
        df = df[df["COB"].isin(selected_cob)]

# ================= UW YEAR (CHECKBOX) =================
uw_col = None
for col in ["UW Year", "UW_YEAR", "UWYear", "Year"]:
    if col in df.columns:
        uw_col = col
        break

if uw_col:
    uw_options = sorted(df[uw_col].dropna().unique().tolist())
    selected_uw = st.sidebar.multiselect("Pilih UW Year", uw_options, default=uw_options)

    if selected_uw:
        df = df[df[uw_col].isin(selected_uw)]

# ===============================
# ZERO OPTION (OPTIONAL)
# ===============================
zero_option = st.sidebar.checkbox("Include Zero Value", value=True)

# ===============================
# FILE NAME
# ===============================
file_name = st.text_input("Nama File Output", value="SOA_Report")

# ===============================
# GENERATE REPORT FUNCTION (SUDAH KAMU PUNYA)
# ===============================
# Pastikan fungsi ini tetap sama seperti punyamu
# def generate_report(df, tipe, zero_option):
#     ...
def generate_report(df, tipe, zero_option):

    if tipe == "QS":
        df['PREMIUM'] = df['qs_ceding']
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

    # ============================
    # LOOP CURRENCY
    # ============================
    for curr, df_curr in grouped.groupby('CURRENCY'):

        # 🔥 FILTER ZERO (LEVEL CURRENCY)
        if zero_option == "Hide Zero Rows":
            df_curr = df_curr[
                ~(
                    (df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']] == 0)
                    .all(axis=1)
                )
            ]

        if df_curr.empty:
            continue

        first_row = True

        # ============================
        # LOOP COB
        # ============================
        for cob, df_cob in df_curr.groupby('COB'):

            # 🔥 FILTER ZERO (LEVEL COB)
            if zero_option == "Hide Zero Rows":
                df_cob = df_cob[
                    ~(
                        (df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']] == 0)
                        .all(axis=1)
                    )
                ]

            if df_cob.empty:
                continue
            first_cob_row = True   # 🔥 TAMBAHAN
            
            # ============================
            # DETAIL ROW
            # ============================
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

            # ============================
            # SUBTOTAL COB
            # ============================
            subtotal = df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
            rows.append(["", f"{cob} TOTAL", "", *subtotal])

        # ============================
        # TOTAL CURRENCY
        # ============================
        total_curr = df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
        rows.append([f"{curr} TOTAL","","", *total_curr])

        # SPASI
        rows.append(["","","","","","",""])

    return pd.DataFrame(
        rows,
        columns=['CURRENCY','COB','UW YEAR','PREMIUM','COMMISSION','CLAIM','AMOUNT']
    )
report_qs = generate_report(df.copy(), "QS", zero_option)
report_sp = generate_report(df.copy(), "SP", zero_option)

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

from openpyxl.styles import Font, Alignment, PatternFill, Border
from openpyxl.drawing.image import Image

header_fill = PatternFill("solid", fgColor="000000")
grey_fill = PatternFill("solid", fgColor="D9D9D9")
white_fill = PatternFill("solid", fgColor="FFFFFF")
no_border = Border()

# ===============================
# WRITE SHEET FUNCTION (SUDAH KAMU PUNYA)
# ===============================


def write_sheet(writer, data, name, tipe, ref):

    data.to_excel(writer, index=False, sheet_name=name, startrow=12)
    ws = writer.sheets[name]

        # ===============================
    # LOGO (optional)
    # ===============================
    try:
        logo = Image("askrindo.jpg")
        logo.height = 60
        logo.width = 140
        ws.add_image(logo, "A1")
    except:
        pass

    # ===============================
    # TITLE
    # ===============================
    ws.merge_cells('A4:G4')
    ws['A4'] = "STATEMENT OF ACCOUNT"
    ws['A4'].font = Font(bold=True, size=14)
    ws['A4'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A5:G5')
    ws['A5'] = f"Ref No. {ref}"
    ws['A5'].font = Font(bold=True)
    ws['A5'].alignment = Alignment(horizontal='center')

    # ===============================
    # HEADER INFO
    # ===============================
    ws['A7'] = "Treaty Year  :"; ws['B7'] = year
    ws['A8'] = "Quarter      :"; ws['B8'] = f"{quarter} {tipe}"
    ws['A9'] = "For Months   :"; ws['B9'] = months_text
    ws['A10'] = "Broker       :"; ws['B10'] = selected_broker
    
    # HEADER TABLE
    for col in "ABCDEFG":
        cell = ws[f"{col}13"]
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.border = no_border

    current_currency = None

    for row in range(14, ws.max_row+1):

        val_curr = ws[f"A{row}"].value
        val_cob  = ws[f"B{row}"].value

        # DEFAULT PUTIH
        for col in "ABCDEFG":
            ws[f"{col}{row}"].fill = white_fill
            ws[f"{col}{row}"].border = no_border

        if all(ws[f"{col}{row}"].value in ["", None] for col in "ABCDEFG"):
            current_currency = None
            continue

        # ALIGNMENT
        ws[f"A{row}"].alignment = Alignment(horizontal='left')
        ws[f"B{row}"].alignment = Alignment(horizontal='left')
        ws[f"C{row}"].alignment = Alignment(horizontal='center')

        # DETECT CURRENCY
        if val_curr not in ["", None] and "TOTAL" not in str(val_curr):
            current_currency = val_curr

        # KOLOM A GREY
        if current_currency:
            ws[f"A{row}"].fill = grey_fill

        # TOTAL CURRENCY FULL GREY
        if val_curr and "TOTAL" in str(val_curr):
            for col in "ABCDEFG":
                ws[f"{col}{row}"].fill = grey_fill
            current_currency = None

        # BOLD
        ws[f"A{row}"].font = Font(bold=True)
        ws[f"B{row}"].font = Font(bold=True)

        if "TOTAL" in str(val_cob) or "TOTAL" in str(val_curr):
            for col in "ABCDEFG":
                ws[f"{col}{row}"].font = Font(bold=True)

        # FORMAT ANGKA
        for col in ['D','E','F','G']:
            ws[f"{col}{row}"].number_format = '#,##0.00;[Red](#,##0.00)'

    # NOTE
    last = ws.max_row + 2
    ws[f"A{last}"] = "Note :"
    ws[f"B{last}"] = note

with pd.ExcelWriter(output, engine='openpyxl') as writer:
    write_sheet(writer, report_qs, "QS Report", "Quota Share", ref_qs)
    write_sheet(writer, report_sp, "SL Report", "Surplus", ref_sp)


# ===============================
# DOWNLOAD ALL BROKER (MULTI SHEET)
# ===============================
if st.button("⬇️ Download All Broker (1 File Multi-Sheet)"):

    output_all = io.BytesIO()

    with pd.ExcelWriter(output_all, engine='openpyxl') as writer:

        if "BROKER" not in df.columns:
            st.error("Kolom BROKER tidak ditemukan di data")
            st.stop()

        brokers = df["BROKER"].dropna().unique()

        for broker in brokers:
            df_broker = df[df["BROKER"] == broker]

            # Generate report
            report_qs = generate_report(df_broker.copy(), "QS", zero_option)
            report_sp = generate_report(df_broker.copy(), "SP", zero_option)

            # Sheet name (limit 31 char)
            sheet_qs = f"QS_{str(broker)}"[:31]
            sheet_sp = f"SP_{str(broker)}"[:31]

            # Write sheet
            write_sheet(writer, report_qs, sheet_qs, "Quota Share", None, broker)
            write_sheet(writer, report_sp, sheet_sp, "Surplus", None, broker)

    st.download_button(
        "📥 Download Excel All Broker",
        data=output_all.getvalue(),
        file_name=f"{file_name}_ALL_BROKER.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
