import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="SOA Processing", layout="wide")

st.title("📊 SOA Processing Tool")
st.write("Upload file SOA dan SOR Summary untuk diproses")

# ===============================
# UPLOAD FILE
# ===============================
file_soa = st.file_uploader("Upload Data SOA (Excel)", type=["xlsx"])
file_sor = st.file_uploader("Upload SOR Summary (Excel)", type=["xlsx"])

# ===============================
# FUNCTION PROCESSING
# ===============================
def process_data(df1, df2):

    # CLEAN DATA PREMI
    df1['QS'] = pd.to_numeric(df1['QS'], errors='coerce').fillna(0)
    df1['SPL'] = pd.to_numeric(df1['SPL'], errors='coerce').fillna(0)
    df1 = df1[~((df1['QS'] == 0) & (df1['SPL'] == 0))]
    df1['COB'] = df1['COB'].astype(str).str.strip().str.upper()

    cols_convert = ['TSI SHARE','OR','QS','SPL']
    for col in cols_convert:
        df1[col] = pd.to_numeric(df1[col], errors='coerce')

    df1['UY-COB'] = df1['UY'].astype(str) + "-" + df1['COB']

    # CLEAN REFERENSI
    df2.columns = df2.columns.str.strip().str.lower()
    df2['uy'] = df2['uy'].astype(str).str.strip()

    for col in ['broker','cob','group']:
        df2[col] = df2[col].astype(str).str.strip().str.upper()

    def percent_to_decimal(x):
        if pd.isna(x):
            return 0
        if isinstance(x, str):
            x = x.replace('%','').strip()
            val = pd.to_numeric(x, errors='coerce')
            return val / 100 if val is not None else 0
        if isinstance(x, (int,float)):
            if x > 1:
                return x / 100
        return x

    df2['sharere']  = df2['sharere'].apply(percent_to_decimal)
    df2['komisiqs'] = df2['komisiqs'].apply(percent_to_decimal)
    df2['komisisp'] = df2['komisisp'].apply(percent_to_decimal)

    df2['cashcall'] = df2['cashcall'].astype(str).str.replace(',', '')
    df2['cashcall'] = pd.to_numeric(df2['cashcall'], errors='coerce')

    # MERGE
    merged = df1.merge(
        df2,
        left_on=['UY','COB'],
        right_on=['uy','cob'],
        how='left'
    )

    # HANDLE MISSING → SPREAD 2023
    found = merged[merged['broker'].notna()].copy()
    missing = merged[merged['broker'].isna()].copy()

    df2_2023 = df2[df2['uy'] == '2023'].copy()

    missing_expanded = missing.drop(columns=df2.columns, errors='ignore').merge(
        df2_2023,
        left_on='COB',
        right_on='cob',
        how='left'
    )

    merged = pd.concat([found, missing_expanded], ignore_index=True)

    # DROP
    merged = merged.drop(columns=['cashcall','cob','uy'], errors='ignore')

    # NUMERIC FINAL
    cols = ['QS','SPL','sharere','komisiqs','komisisp']
    for c in cols:
        merged[c] = pd.to_numeric(merged[c], errors='coerce').fillna(0)

    # HITUNG
    merged['qs_ceding'] = merged['QS'] * merged['sharere']
    merged['komisi_qs'] = merged['qs_ceding'] * merged['komisiqs']

    merged['sp_ceding'] = merged['SPL'] * merged['sharere']
    merged['komisi_sp'] = merged['sp_ceding'] * merged['komisisp']

    return merged


# ===============================
# MAIN PROCESS
# ===============================

def add_total_row(df):
    numeric_cols = df.select_dtypes(include='number').columns
    total = df[numeric_cols].sum()
    
    total_df = pd.DataFrame([total])
    total_df.index = ['TOTAL']
    
    return total_df

#============================================
if file_soa and file_sor:

    df1 = pd.read_excel(file_soa)
    df2 = pd.read_excel(file_sor)
    
    #=====SOA
    st.subheader("📄 Data SOA (Original)")
    st.dataframe(df1, use_container_width=True)
    
    #======SOR
    st.subheader("📄 Data SOR (Original)")
    st.dataframe(df2, use_container_width=True)

    # PROCESS
    result = process_data(df1.copy(), df2.copy())

    st.subheader("✅ Hasil Setelah Proses")
    st.dataframe(result, use_container_width=True)

    #===TOTAL================
    st.subheader("🔢 Total Data SOA")
    total_soa = add_total_row(df1)
    st.dataframe(total_soa, use_container_width=True)
    
    st.subheader("🔢 Total Hasil Output")
    total_result = add_total_row(result)
    st.dataframe(total_result, use_container_width=True)

    # ===============================
    # DOWNLOAD EXCEL
    # ===============================
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, index=False, sheet_name='Result')

    # gunakan nama dari user
    final_filename = file_name_input.strip()
    
    # fallback kalau kosong
    if final_filename == "":
        final_filename = "SOA_Result"
    
    st.download_button(
        label="⬇️ Download Hasil Excel",
        data=output.getvalue(),
        file_name=f"{final_filename}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Silakan upload kedua file terlebih dahulu")
