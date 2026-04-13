import streamlit as st
import pandas as pd
import io

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_bg(cell, color="FFFFFF"):
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def remove_table_borders(table):
    tbl = table._element
    tblPr = tbl.tblPr

    borders = OxmlElement('w:tblBorders')

    for edge in ('left', 'right'):
        elem = OxmlElement(f'w:{edge}')
        elem.set(qn('w:val'), 'nil')
        borders.append(elem)

    tblPr.append(borders)

def format_number(val):
    try:
        val = float(val)
        if val < 0:
            return f"({abs(val):,.2f})", True
        else:
            return f"{val:,.2f}", False
    except:
        return str(val), False


def export_to_word_clean(df, broker_loop, file_name):

    doc = Document()
    from docx.shared import Mm

    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    
    # margin kecil biar muat banyak
    section.top_margin = Mm(10)
    section.bottom_margin = Mm(10)
    section.left_margin = Mm(10)
    section.right_margin = Mm(10)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(8)

    for idx, broker in enumerate(broker_loop):
        ref_number = start_number + idx
        roman_quarter = to_roman(quarter)
        
        ref_auto = f"{ref_number}/UDWR/{roman_quarter}/{year}"
        df_broker = df[df['BROKER'] == broker]

        report = generate_report(df_broker.copy(), "QS", zero_option)

        # =========================
        # HEADER
        # =========================
        try:
            doc.add_picture("askrindo.jpg", width=Inches(1.5))
        except:
            pass

        title = doc.add_paragraph("STATEMENT OF ACCOUNT")
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].bold = True

        # 🔥 REF NO DI TENGAH (TEPAT DI BAWAH TITLE)
        p = doc.add_paragraph(f"Ref No : {ref_auto}")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        
        p = doc.add_paragraph(f"Treaty Year : {year}")
        p.paragraph_format.space_after = Pt(0)
        
        p = doc.add_paragraph(f"Quarter : {quarter}")
        p.paragraph_format.space_after = Pt(0)
        
        p = doc.add_paragraph(f"For Months : {months_text}")
        p.paragraph_format.space_after = Pt(2)

        doc.add_paragraph("")

        # =========================
        # TABLE HEADER
        # =========================
        table = doc.add_table(rows=1, cols=8)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False 
        
        # 🔥 HAPUS GRID
        remove_table_borders(table)
        
        # 🔥 SET LEBAR KOLOM
        col_widths = [1, 2.2, 0.8, 1.2, 1.2, 1.2, 1.2, 1.2]

        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(width)

        headers = ['CURRENCY','COB','UY','PREMIUM','COMMISSION','CLAIM','RECOVERY','AMOUNT']

        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = h
        
            # 🔥 BACKGROUND HITAM
            set_cell_bg(cell, "000000")
            def set_row_border(cells):
                from docx.oxml import OxmlElement
                from docx.oxml.ns import qn
            
                for cell in cells:
                    tc = cell._element
                    tcPr = tc.get_or_add_tcPr()
            
                    tcBorders = OxmlElement('w:tcBorders')
            
                    for border_name in ['top', 'bottom']:
                        border = OxmlElement(f'w:{border_name}')
                        border.set(qn('w:val'), 'single')
                        border.set(qn('w:sz'), '12')  # 🔥 tebal
                        border.set(qn('w:space'), '0')
                        border.set(qn('w:color'), '000000')
                        tcBorders.append(border)
            
                    tcPr.append(tcBorders)
        
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.runs[0]
            run.bold = True
            run.font.color.rgb = RGBColor(255,255,255)

        current_currency = None

        # =========================
        # LOOP DATA (🔥 RAPiH)
        # =========================
        for _, row in report.iterrows():

            values = list(row)

            cells = table.add_row().cells

            for i, val in enumerate(values):

                text, is_negative = format_number(val) if i >= 3 else (str(val), False)

                cells[i].text = text

                para = cells[i].paragraphs[0]

                # alignment
                if i >= 3:
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif i == 2:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT

                # warna merah untuk negatif
                if is_negative:
                    para.runs[0].font.color.rgb = RGBColor(255, 0, 0)

                # bold untuk TOTAL
                row_text = " ".join([str(v) for v in values])

                # 🔥 DETECT TYPE TOTAL
                is_currency_total = "TOTAL" in str(values[0])   # contoh: USD TOTAL
                is_cob_total = "TOTAL" in str(values[1])        # contoh: FIRE TOTAL
                is_grand_total = "GRAND TOTAL" in row_text
                
                # =========================
                # APPLY STYLE PER ROW
                # =========================
                if is_currency_total or is_grand_total:
                
                    # 🔥 background abu hanya untuk currency total
                    for c in cells:
                        set_cell_bg(c, "D9D9D9")
                
                    # 🔥 bold
                    for c in cells:
                        for p in c.paragraphs:
                            for r in p.runs:
                                r.bold = True
                
                    # 🔥 garis tebal
                    set_row_border(cells)
                
                elif is_cob_total:
                
                    # 🔥 hanya bold + garis (tanpa abu)
                    for c in cells:
                        for p in c.paragraphs:
                            for r in p.runs:
                                r.bold = True
                
                    set_row_border(cells)        
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(0)

        # # =========================
        # # NOTE
        # # =========================
        # doc.add_paragraph("")
        # doc.add_paragraph(f"Note : {note}")

        # TTD (KIRI + KANAN)
        # =========================
        doc.add_paragraph("")
        
        # 🔥 BARIS ATAS (KIRI & KANAN)
        p = doc.add_paragraph()
        
        run_left = p.add_run("Agreed and approved by Reinsurer")
        run_left.font.size = Pt(8)
        
        # kasih jarak manual biar ke kanan
        p.add_run("\t\t\t\t\t\t\t\t\t\t")
        report_date_str = report_date.strftime("%d %B %Y")
        run_right = p.add_run(f"Jakarta, {report_date_str}")
        run_right.font.size = Pt(8)
        
        # 🔥 BARIS KEDUA (KIRI & KANAN)
        p2 = doc.add_paragraph()
        
        run_left2 = p2.add_run(f"{broker}")
        run_left2.bold = True
        
        p2.add_run("\t\t\t\t\t\t\t\t\t\t\t\t")
        
        run_right2 = p2.add_run("PT. Asuransi Kredit Indonesia")
        run_right2 = p2.add_run("Underwriting & Reinsurance Division")
        # 🔥 SPASI TTD
        doc.add_paragraph("")
        doc.add_paragraph("")
        
        # 🔥 NAMA (KANAN)
        ttd3 = doc.add_paragraph("Budi Santoso AI\nDivision Head")
        ttd3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # doc.add_paragraph("")
        # ttd = doc.add_paragraph("Jakarta, January 2026")
        # ttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # ttd2 = doc.add_paragraph("PT. Asuransi Kredit Indonesia\nUnderwriting & Reinsurance Division")
        # ttd2.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # doc.add_paragraph("")
        # ttd3 = doc.add_paragraph("Budi Santoso AI\nDivision Head")
        # ttd3.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # =========================
        # PAGE BREAK
        # =========================
        if idx < len(broker_loop) - 1:
            doc.add_page_break()

    import io
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return file_stream

def format_quarter_text(q):
    mapping = {
        "SP": "Surplus",
        "QS": "Quota Share"
    }
    return mapping.get(q, q)

st.set_page_config(page_title="SOA Financial Report", layout="wide")
st.title("📑 SOA Financial Report Generator")

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
# ===============================
# MAPPING PRODUCT → COB
# ===============================
# if 'product' in df.columns:
#     df['cob'] = df['product']
    
mapping = {
    'prod':'PROD',
    'product':'COB',
    'cob':'COB',
    'uy':'UY',

    'valuta':'CURRENCY',
    'curr':'CURRENCY',
    'currency':'CURRENCY',

    'broker':'BROKER',

    # 🔥 SESUAI FILE ASLI
    'premi_panel_qs':'PREMI_PANEL_QS',
    'premi_panel_sp':'PREMI_PANEL_SP',

    'komisi_panel_qs':'KOMISI_PANEL_QS',
    'komisi_panel_sp':'KOMISI_PANEL_SP',

    'klaim_panel_qs':'KLAIM_PANEL_QS',
    'klaim_panel_sp':'KLAIM_PANEL_SP',

    # ❗ INI YANG TADI SALAH
    'recoveries_panel_qs':'RECOVERY_PANEL_QS',
    'recoveries_panel_sp':'RECOVERY_PANEL_SP'
}
df = df.rename(columns=mapping)


# ===============================
# FIX TEXT
# ===============================
for col in ['CURRENCY', 'COB', 'BROKER']:
    if col in df.columns:
        df[col] = (
            df[col]
            .fillna("")          # 🔥 hindari NaN error
            .astype(str)         # ubah ke string
        )
        df[col] = df[col].str.strip().str.upper()

df = df[(df['CURRENCY'] != "") & (df['CURRENCY'].notna())]

# ===============================
# CLEAN UY (🔥 TARUH DI SINI)
# ===============================
def clean_uy(x):
    x = str(x).strip()

    # handle format: 2018/2019 atau 2018-2019
    if "/" in x or "-" in x:
        sep = "/" if "/" in x else "-"
        try:
            return int(x.split(sep)[0])  # ambil tahun depan
        except:
            return None

    # handle normal number
    try:
        return int(float(x))
    except:
        return None
# ===============================
# FILTER COB (CHECKBOX)
# ===============================
st.subheader("Pilih COB")

if "COB" in df.columns:
    cob_list = sorted(df["COB"].dropna().unique().tolist())

    selected_cob = []

    all_cob = st.checkbox("ALL COB", value=True)

    if all_cob:
        selected_cob = cob_list
    else:
        cols = st.columns(3)  # biar rapi grid
        for i, cob in enumerate(cob_list):
            if cols[i % 3].checkbox(cob):
                selected_cob.append(cob)

    if selected_cob:
        df = df[df["COB"].isin(selected_cob)]
    else:
        st.warning("Pilih minimal 1 COB")
        st.stop()

# ===============================
# FILTER UW YEAR (CHECKBOX)
# ===============================
st.subheader("Pilih UW Year")

if "UY" in df.columns:
    uy_list = sorted(df["UY"].dropna().unique().tolist())

    selected_uy = []

    all_uy = st.checkbox("ALL UW YEAR", value=True)

    if all_uy:
        selected_uy = uy_list
    else:
        cols = st.columns(4)
        for i, uy in enumerate(uy_list):
            if cols[i % 4].checkbox(str(uy)):
                selected_uy.append(uy)

    if selected_uy:
        df = df[df["UY"].isin(selected_uy)]
    else:
        st.warning("Pilih minimal 1 UW Year")
        st.stop()
        
# ===============================
# FILTER
# ===============================
broker_list = df['BROKER'].dropna().unique().tolist()
broker_list.insert(0, "ALL")
selected_broker = st.selectbox("Pilih Broker", broker_list)

if selected_broker != "ALL":
    df = df[df['BROKER'] == selected_broker]

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
# USE BULAN & TAHUN LANGSUNG
# ===============================
df['MONTH'] = pd.to_numeric(df['bulan'], errors='coerce')
df['YEAR']  = pd.to_numeric(df['tahun'], errors='coerce')

df = df.dropna(subset=['MONTH','YEAR'])

# ===============================
# DATE INFO
# ===============================
month_full = {
    1:'January',2:'February',3:'March',4:'April',
    5:'May',6:'June',7:'July',8:'August',
    9:'September',10:'October',11:'November',12:'December'
}

min_m = int(df['MONTH'].min())
max_m = int(df['MONTH'].max())
year = int(df['YEAR'].mode()[0])
months_text = f"{month_full[min_m]} {year} - {month_full[max_m]} {year}"
# month_map = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
#              7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

# min_m = int(df['MONTH'].min())
# max_m = int(df['MONTH'].max())
# year = int(df['YEAR'].mode()[0])

# months_text = f"{month_map[min_m]} - {month_map[max_m]} {year}"
def to_roman(q):
    mapping = {
        "I": "I",
        "II": "II",
        "III": "III",
        "IV": "IV"
    }
    return mapping.get(q, q)
    
def get_quarter(m):
    return ["I","II","III","IV"][(m-1)//3]

quarter = get_quarter(max_m)

required_cols = ['CURRENCY','COB','UY']
missing = [c for c in required_cols if c not in df.columns]

if missing:
    st.error(f"Kolom wajib tidak ditemukan: {missing}")
    st.stop()

# ===============================
# CLEAN NUMERIC COLUMNS
# ===============================
numeric_cols = [
    'PREMI_PANEL_QS','KOMISI_PANEL_QS','KLAIM_PANEL_QS','RECOVERY_PANEL_QS',
    'PREMI_PANEL_SP','KOMISI_PANEL_SP','KLAIM_PANEL_SP','RECOVERY_PANEL_SP'
]

for col in numeric_cols:
    if col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)  # hapus koma
            .str.replace(" ", "", regex=False)  # hapus spasi
        )
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
# ===============================
# GENERATE REPORT
# ===============================
def generate_report(df, tipe, zero_option):

    # ===============================
    # AMBIL DATA FINANCIAL
    # ===============================
    cols = [
        'PREMI_PANEL_QS','KOMISI_PANEL_QS','KLAIM_PANEL_QS','RECOVERY_PANEL_QS',
        'PREMI_PANEL_SP','KOMISI_PANEL_SP','KLAIM_PANEL_SP','RECOVERY_PANEL_SP'
    ]
    
    for col in cols:
        df[col] = pd.to_numeric(df.get(col, 0), errors='coerce').fillna(0)
    
    # ===============================
    # PILIH TIPE
    # ===============================
    if tipe == "QS":
        df['PREMIUM'] = df['PREMI_PANEL_QS']
        df['COMMISSION'] = df['KOMISI_PANEL_QS']
        df['CLAIM'] = df['KLAIM_PANEL_QS']
        df['RECOVERY'] = df['RECOVERY_PANEL_QS']
    else:
        df['PREMIUM'] = df['PREMI_PANEL_SP']
        df['COMMISSION'] = df['KOMISI_PANEL_SP']
        df['CLAIM'] = df['KLAIM_PANEL_SP']  
        df['RECOVERY'] = df['RECOVERY_PANEL_SP']
        
    df['AMOUNT'] = df['PREMIUM'] - df['COMMISSION'] - df['CLAIM'] +df['RECOVERY']

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
                    (df_curr[['PREMIUM','COMMISSION','CLAIM','RECOVERY','AMOUNT']] == 0)
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
                    r['RECOVERY'],
                    r['AMOUNT']
                ])
                first_row = False
                first_cob_row = False

            # ============================
            # SUBTOTAL COB
            # ============================
            subtotal = df_cob[['PREMIUM','COMMISSION','CLAIM','RECOVERY','AMOUNT']].sum()
            rows.append(["", f"{cob} TOTAL", "", *subtotal])

        # ============================
        # TOTAL CURRENCY
        # ============================
        total_curr = df_curr[['PREMIUM','COMMISSION','CLAIM','RECOVERY','AMOUNT']].sum()
        rows.append([f"{curr} TOTAL","","", *total_curr])

        # SPASI
        rows.append(["","","","","","",""])
    # ============================
    # GRAND TOTAL
    # ============================
    grand_total = grouped[['PREMIUM','COMMISSION','CLAIM','RECOVERY','AMOUNT']].sum()
            
    rows.append(["", "GRAND TOTAL", "", *grand_total])
            
    # spasi akhir
    rows.append(["","","","","","",""])        

    return pd.DataFrame(
        rows,
        columns=['CURRENCY','COB','UW YEAR','PREMIUM','COMMISSION','CLAIM','RECOVERY','AMOUNT']
    )
report_qs = generate_report(df.copy(), "QS", zero_option)
report_sp = generate_report(df.copy(), "SP", zero_option)

st.dataframe(report_qs)

# ===============================
# INPUT
# ===============================
start_number = st.number_input("Nomor Awal Ref No", value=81, step=1)
# ref_qs = st.text_input("Ref No QS")
# ref_sp = st.text_input("Ref No SPL")
# note = st.text_area("Note")
import datetime
report_date = st.date_input("Pilih Tanggal", datetime.date.today())
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

def write_combined_sheet(writer, qs_data, sp_data, sheet_name, broker, ref_qs, ref_sp):

    qs_start = 12
    sp_start = qs_start + len(qs_data) + 15
    title_row = sp_start - 6

    qs_data.to_excel(writer, index=False, sheet_name=sheet_name, startrow=qs_start)
    sp_data.to_excel(writer, index=False, sheet_name=sheet_name, startrow=sp_start)

    ws = writer.sheets[sheet_name]

    # ===============================
    # LOGO QS
    # ===============================
    try:
        logo = Image("askrindo.jpg")
        logo.height = 60
        logo.width = 140
        ws.add_image(logo, "A1")
    except:
        pass

    # ===============================
    # TITLE QS
    # ===============================
    ws.merge_cells('A4:G4')
    ws['A4'] = "STATEMENT OF ACCOUNT - QS"
    ws['A4'].font = Font(bold=True, size=14)
    ws['A4'].alignment = Alignment(horizontal='center')

    ws['A7'] = "Broker :"
    ws['B7'] = broker
    ws.merge_cells('A5:G5')
    ws['A5'] = f"Ref No. {ref_qs}"
    ws['A5'].font = Font(bold=True)
    ws['A5'].alignment = Alignment(horizontal='center')
    
    ws['A7'] = "Treaty Year  :"; ws['B7'] = year
    ws['B8'] = f"{quarter} {format_quarter_text('QS')}"
    ws[f"B{title_row+4}"] = f"{quarter} {format_quarter_text('SP')}"
    

    # ===============================
    # TITLE SP
    # ===============================
   

    ws.merge_cells(f'A{title_row}:G{title_row}')
    ws[f'A{title_row}'] = "STATEMENT OF ACCOUNT - SP"
    ws[f'A{title_row}'].font = Font(bold=True, size=14)
    ws[f'A{title_row}'].alignment = Alignment(horizontal='center')

    ws[f"A{title_row+2}"] = "Broker :"
    ws[f"B{title_row+2}"] = broker
    ws.merge_cells(f'A{title_row+1}:G{title_row+1}')
    ws[f'A{title_row+1}'] = f"Ref No. {ref_sp}"
    ws[f'A{title_row+1}'].font = Font(bold=True)
    ws[f'A{title_row+1}'].alignment = Alignment(horizontal='center')
    
    ws[f"A{title_row+3}"] = "Treaty Year  :"; ws[f"B{title_row+3}"] = year
    ws[f"A{title_row+4}"] = "Quarter      :"; ws[f"B{title_row+4}"] = f"{quarter} SP"
    ws[f"A{title_row+5}"] = "For Months   :"; ws[f"B{title_row+5}"] = months_text
    ws[f"A{title_row+6}"] = "Broker       :"; ws[f"B{title_row+6}"] = broker

    # ===============================
    # LOGO SP
    # ===============================
    try:
        logo2 = Image("askrindo.jpg")
        logo2.height = 60
        logo2.width = 140
        ws.add_image(logo2, f"A{sp_start-10}")
    except:
        pass

    # ===============================
    # FUNCTION STYLING
    # ===============================
    def apply_style(start_row, end_row):

        # HEADER TABLE
        for col in "ABCDEFGH":
            cell = ws[f"{col}{start_row+1}"]
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.border = no_border

        current_currency = None

        for row in range(start_row+2, end_row+1):

            val_curr = ws[f"A{row}"].value
            val_cob  = ws[f"B{row}"].value

            # default putih
            for col in "ABCDEFGH":
                ws[f"{col}{row}"].fill = white_fill
                ws[f"{col}{row}"].border = no_border

            # skip baris kosong
            if all(ws[f"{col}{row}"].value in ["", None] for col in "ABCDEFGH"):
                current_currency = None
                continue

            # alignment
            ws[f"A{row}"].alignment = Alignment(horizontal='left')
            ws[f"B{row}"].alignment = Alignment(horizontal='left')
            ws[f"C{row}"].alignment = Alignment(horizontal='center')

            # detect currency
            if val_curr not in ["", None] and "TOTAL" not in str(val_curr):
                current_currency = val_curr

            # kolom A grey
            if current_currency:
                ws[f"A{row}"].fill = grey_fill

            # TOTAL ROW FULL GREY + BOLD
            if (val_curr and "TOTAL" in str(val_curr)) or (val_cob and "TOTAL" in str(val_cob)):
                for col in "ABCDEFGH":
                    ws[f"{col}{row}"].fill = grey_fill
                    ws[f"{col}{row}"].font = Font(bold=True)
                current_currency = None

            # bold kolom A & B
            ws[f"A{row}"].font = Font(bold=True)
            ws[f"B{row}"].font = Font(bold=True)

            # format angka (kurung merah otomatis)
            for col in ['D','E','F','G','H']:
                ws[f"{col}{row}"].number_format = '#,##0.00;[Red](#,##0.00)'

    # ===============================
    # APPLY STYLE QS
    # ===============================
    qs_end = qs_start + len(qs_data) + 1
    apply_style(qs_start, qs_end)

    # ===============================
    # APPLY STYLE SP
    # ===============================
    sp_end = sp_start + len(sp_data) + 1
    apply_style(sp_start, sp_end)

    # # ===============================
    # # NOTE
    # # ===============================
    # last = ws.max_row + 2
    # ws[f"A{last}"] = "Note :"
    # ws[f"B{last}"] = note
    
with pd.ExcelWriter(output, engine='openpyxl') as writer:

    if selected_broker == "ALL":
        broker_loop = df['BROKER'].dropna().unique()
    else:
        broker_loop = [selected_broker]

    for broker in broker_loop:

        df_broker = df[df['BROKER'] == broker]

        # 🔥 INI WAJIB PER BROKER
        report_qs = generate_report(df_broker.copy(), "QS", zero_option)
        report_sp = generate_report(df_broker.copy(), "SP", zero_option)

        write_combined_sheet(
            writer,
            report_qs,
            report_sp,
            sheet_name=str(broker)[:31],
            broker=broker,
            ref_qs=ref_qs,
            ref_sp=ref_sp
        )
        
st.download_button(
    "⬇️ Download Report",
    data=output.getvalue(),
    file_name=f"{file_name}.xlsx"
)

# ===============================
# EXPORT WORD
# ===============================
if st.button("📄 Export Word (Rapi)"):
    
    if selected_broker == "ALL":
        broker_loop = df['BROKER'].dropna().unique()
    else:
        broker_loop = [selected_broker]

    file_stream = export_to_word_clean(df, broker_loop, file_name)

    st.download_button(
        "⬇️ Download Word",
        file_stream,
        file_name=f"{file_name}.docx"
    )
