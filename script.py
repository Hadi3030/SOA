import streamlit as st
import pandas as pd
import io
from PIL import Image

st.set_page_config(page_title="SOA Processing", layout="wide")

# ===============================
# LOAD LOGO
# ===============================
logo_path = "askrindo.jpg"

# ===============================
# SIDEBAR MENU
# ===============================
st.sidebar.title("📌 Menu")

mode = st.sidebar.radio(
    "Pilih Analisis",
    ["Spreading Data (SOA Processing)", "Laporan SOA (SOA Report)"]
)

# ===============================
# SESSION STATE
# ===============================
if "result_data" not in st.session_state:
    st.session_state["result_data"] = None

# ===============================
# FUNCTION: STANDARDISASI KOLOM
# ===============================
def standardize_columns(df):
    df.columns = df.columns.str.strip().str.upper()
    return df

# ===============================
# FUNCTION PROCESSING
# ===============================
def process_data(df1, df2):

    df1 = standardize_columns(df1)
    df2 = standardize_columns(df2)

    df1['QS'] = pd.to_numeric(df1.get('QS', 0), errors='coerce').fillna(0)
    df1['SPL'] = pd.to_numeric(df1.get('SPL', 0), errors='coerce').fillna(0)

    df1 = df1[~((df1['QS'] == 0) & (df1['SPL'] == 0))]

    df1['COB'] = df1['COB'].astype(str).str.strip().str.upper()

    cols_convert = ['TSI SHARE','OR','QS','SPL']
    for col in cols_convert:
        if col in df1.columns:
            df1[col] = pd.to_numeric(df1[col], errors='coerce')

    df1['UY-COB'] = df1['UY'].astype(str) + "-" + df1['COB']

    def percent_to_decimal(x):
        if pd.isna(x):
            return 0
        if isinstance(x, str):
            x = x.replace('%','').strip()
            val = pd.to_numeric(x, errors='coerce')
            return val / 100 if val is not None else 0
        if isinstance(x, (int,float)) and x > 1:
            return x / 100
        return x

    for col in ['SHARERE','KOMISIQS','KOMISISP']:
        if col in df2.columns:
            df2[col] = df2[col].apply(percent_to_decimal)

    if 'CASHCALL' in df2.columns:
        df2['CASHCALL'] = df2['CASHCALL'].astype(str).str.replace(',', '')
        df2['CASHCALL'] = pd.to_numeric(df2['CASHCALL'], errors='coerce')

    merged = df1.merge(
        df2,
        left_on=['UY','COB'],
        right_on=['UY','COB'],
        how='left'
    )

    found = merged[merged['BROKER'].notna()].copy()
    missing = merged[merged['BROKER'].isna()].copy()

    df2_2023 = df2[df2['UY'] == '2023'].copy()

    missing_expanded = missing.drop(columns=df2.columns, errors='ignore').merge(
        df2_2023,
        on='COB',
        how='left'
    )

    merged = pd.concat([found, missing_expanded], ignore_index=True)

    merged = merged.drop(columns=['CASHCALL'], errors='ignore')

    for c in ['QS','SPL','SHARERE','KOMISIQS','KOMISISP']:
        if c in merged.columns:
            merged[c] = pd.to_numeric(merged[c], errors='coerce').fillna(0)

    merged['QS_CEDING'] = merged['QS'] * merged['SHARERE']
    merged['KOMISI_QS'] = merged['QS_CEDING'] * merged['KOMISIQS']

    merged['SP_CEDING'] = merged['SPL'] * merged['SHARERE']
    merged['KOMISI_SP'] = merged['SP_CEDING'] * merged['KOMISISP']

    return merged

# ===============================
# FUNCTION HEADER EXCEL
# ===============================
def write_header_excel(writer, sheet_name, ref_no, treaty_year, quarter, for_months, remarks):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # LOGO KIRI ATAS
    worksheet.insert_image('A1', logo_path, {'x_scale': 0.5, 'y_scale': 0.5})

    # FORMAT
    bold = workbook.add_format({'bold': True, 'font_size': 12})
    normal = workbook.add_format({'font_size': 10})

    # JUDUL
    worksheet.write('D2', 'STATEMENT OF ACCOUNT', bold)
    worksheet.write('D3', f"Ref No. {ref_no}", normal)

    # DETAIL
    worksheet.write('A5', f"Treaty Year : {treaty_year}", normal)
    worksheet.write('A6', f"Quarter : {quarter}", normal)
    worksheet.write('A7', f"For Months : {for_months}", normal)
    worksheet.write('A8', f"Remarks : {remarks}", normal)

    # SPACING BIAR MIRIP TEMPLATE
    worksheet.set_row(1, 25)
    worksheet.set_column('A:A', 30)
    worksheet.set_column('D:D', 40)

# ===============================
# FUNCTION REPORT
# ===============================
def generate_report(df):

    df = standardize_columns(df)

    if 'CURRENCY' not in df.columns:
        df['CURRENCY'] = 'IDR'

    if 'QS' not in df.columns or 'SPL' not in df.columns:
        st.error("Kolom QS / SPL tidak ditemukan")
        st.stop()

    df['PREMIUM'] = df['QS'] + df['SPL']
    df['COMMISSION'] = df['KOMISI_QS'] + df['KOMISI_SP']
    df['CLAIM'] = 0
    df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION']

    grouped = df.groupby(['CURRENCY','COB','UY']).agg({
        'PREMIUM':'sum',
        'COMMISSION':'sum',
        'CLAIM':'sum',
        'AMOUNT':'sum'
    }).reset_index()

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

            total = df_cob[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
            final_rows.append(["", f"{cob} TOTAL", "", *total])

        total_curr = df_curr[['PREMIUM','COMMISSION','CLAIM','AMOUNT']].sum()
        final_rows.append([f"{curr} TOTAL","","", *total_curr])

    return pd.DataFrame(final_rows,
        columns=['Currency','COB','UY','Premium','Commission','Claim','Amount'])

# ===============================
# MODE 1 (TIDAK DIUBAH)
# ===============================
if mode == "Spreading Data (SOA Processing)":
    st.title("📊 SOA Processing")

    file_soa = st.file_uploader("Upload Data SOA", type=["xlsx"])
    file_sor = st.file_uploader("Upload SOR Summary", type=["xlsx"])

    if file_soa and file_sor:
        df1 = pd.read_excel(file_soa)
        df2 = pd.read_excel(file_sor)

        st.dataframe(df1)

        result = process_data(df1.copy(), df2.copy())
        st.dataframe(result)

        st.session_state["result_data"] = result

        st.subheader("Total SOA")
        st.dataframe(df1.select_dtypes(include='number').sum())

        st.subheader("Total Output")
        st.dataframe(result.select_dtypes(include='number').sum())

        name = st.text_input("Nama file", "SOA_Result")

        output = io.BytesIO()
        result.to_excel(output, index=False)

        st.download_button("Download",
            data=output.getvalue(),
            file_name=f"{name}.xlsx"
        )

# ===============================
# MODE 2 (REPORT - SUDAH FIX)
# ===============================
else:

    st.title("📑 SOA Report")

    st.markdown("### 🧾 Informasi Header")

    col1, col2 = st.columns(2)

    with col1:
        ref_no = st.text_input("Ref No", "....../DUWR/..../20..")
        treaty_year = st.text_input("Treaty Year", "2026")
        quarter = st.text_input("Quarter", "IV Quota Share")

    with col2:
        for_months = st.text_input("For Months", "Oct - Dec 2025")
        remarks = st.text_input("Remarks", "-")

    source = st.radio(
        "Sumber Data",
        ["Gunakan hasil sebelumnya", "Upload ulang"]
    )

    if source == "Gunakan hasil sebelumnya":
        if st.session_state["result_data"] is None:
            st.warning("Proses data dulu")
            st.stop()
        df = st.session_state["result_data"]

    else:
        file = st.file_uploader("Upload hasil", type=["xlsx"])
        if not file:
            st.stop()
        df = pd.read_excel(file)

    st.dataframe(df)

    report = generate_report(df)

    st.subheader("Report")
    st.dataframe(report)

    name = st.text_input("Nama file report", "SOA_Report")

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        report.to_excel(writer, index=False, startrow=10, sheet_name='Report')

        write_header_excel(
            writer,
            'Report',
            ref_no,
            treaty_year,
            quarter,
            for_months,
            remarks
        )

    st.download_button("Download Report",
        data=output.getvalue(),
        file_name=f"{name}.xlsx"
    )
