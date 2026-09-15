import streamlit as st
import pandas as pd
import datetime
import io
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Set Page Config
st.set_page_config(
    page_title="Sistem Serah Terima Kasir - RS Adhyaksa Jatim",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header { font-size: 24px; font-weight: bold; color: #1e4d2b; text-align: center; margin-bottom: 5px; }
    .sub-header { font-size: 15px; color: #444; text-align: center; margin-bottom: 20px; }
    .stButton>button { width: 100%; background-color: #1e4d2b; color: white; font-weight: bold; }
    .stAlert { padding: 8px 15px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏥 SISTEM INTEGRASI SERAH TERIMA CLOSING KASIR</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Rumah Sakit Adhyaksa Jawa Timur</div>', unsafe_allow_html=True)

# Sidebar: File Upload & Configuration
st.sidebar.header("📂 1. Upload Tarikan File SIMRS")
uploaded_file = st.sidebar.file_uploader("Unggah File SIMRS (.xlsx / .csv)", type=["xlsx", "csv", "xls"])

df = None
file_parsed = False

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        file_parsed = True
        st.sidebar.success("✅ File SIMRS berhasil diuraikan!")
    except Exception as e:
        st.sidebar.error(f"❌ Gagal membaca file: {e}")
else:
    st.sidebar.info("📌 Unggah file laporan SIMRS untuk merekap transaksi & biaya admin EDC secara otomatis.")

# ==========================================
# MODUL UTAMA FORM SERAH TERIMA KASIR
# ==========================================
with st.container():
    col1, col2, col3 = st.columns([1.1, 1, 1])
    
    with col1:
        st.subheader("📋 Informasi Shift & Tanggal")
        tgl_shift = st.date_input("Tanggal Shift", value=datetime.date.today())
        shift_opt = st.selectbox("Shift Operasional", [
            "PAGI (07.00 - 14.00 WIB)",
            "SIANG (14.00 - 21.00 WIB)",
            "MALAM (21.00 - 07.00 WIB)"
        ])
        petugas_lama = st.text_input("Petugas Shift Lama (Menyerahkan)", "CHORI CHOIRUNNISA'")
        petugas_baru = st.text_input("Petugas Shift Baru (Menerima)", "ABDUL JALIL SANTRI AJI")
        pj_kasir = st.text_input("Penanggung Jawab Kasir", "")

    # Variabel Default Rekap
    penerimaan_tunai = 0.0
    total_non_tunai_bruto = 0.0
    total_biaya_admin = 0.0
    df_grouped_final = pd.DataFrame()
    df_display_clean = pd.DataFrame()

    # Processing SIMRS Data
    if file_parsed and df is not None:
        df.columns = [str(c).strip() for c in df.columns]
        
        cols_map = {str(c).strip().lower(): c for c in df.columns}
        col_tgl = cols_map.get('tanggal') or cols_map.get('no. tanggal') or cols_map.get('tgl_transaksi')
        col_notx = cols_map.get('no.rawat/no.nota') or cols_map.get('no. rawat/no. nota') or cols_map.get('no_rawat') or cols_map.get('no. transaksi')
        col_pasien = cols_map.get('nama pasien') or cols_map.get('pasien')
        col_jenis = cols_map.get('jenis/cara bayar') or cols_map.get('jenis pembayaran') or cols_map.get('metode pembayaran')
        col_bersih = cols_map.get('pendapatan bersih') or cols_map.get('total transaksi') or cols_map.get('nominal')

        # Deteksi kolom biaya admin secara fleksibel seperti versi sebelumnya
        col_admin = None
        for c_low, c_orig in cols_map.items():
            if 'admin' in c_low or 'fee' in c_low or 'edc' in c_low or 'qris' in c_low:
                col_admin = c_orig
                break

        def clean_numeric(val):
            if pd.isna(val): return 0.0
            val_str = str(val).replace('.', '').replace(',', '.')
            try:
                return float(val_str)
            except:
                return 0.0

        if col_bersih and col_bersih in df.columns:
            df['clean_bersih'] = df[col_bersih].apply(clean_numeric)
        else:
            df['clean_bersih'] = 0.0
            
        if col_admin and col_admin in df.columns:
            df['clean_admin'] = df[col_admin].apply(clean_numeric)
        else:
            df['clean_admin'] = 0.0

        if col_notx and col_notx in df.columns:
            df_valid = df[df[col_notx].astype(str).str.strip().ne('') & df[col_notx].notna()].copy()
            footer_keywords = ['total', 'jumlah', 'grand']
            for kw in footer_keywords:
                df_valid = df_valid[~df_valid[col_notx].astype(str).str.lower().str.contains(kw, na=False)]
        else:
            df_valid = df.copy()

        if not df_valid.empty:
            group_col = col_notx if col_notx and col_notx in df_valid.columns else df_valid.columns[0]
            
            agg_dict = {
                'clean_bersih': 'first',
                'clean_admin': 'sum'
            }
            if col_tgl and col_tgl in df_valid.columns:
                agg_dict[col_tgl] = 'first'
            if col_pasien and col_pasien in df_valid.columns:
                agg_dict[col_pasien] = 'first'
            if col_jenis and col_jenis in df_valid.columns:
                agg_dict[col_jenis] = 'first'

            df_grouped_final = df_valid.groupby(group_col, as_index=False).agg(agg_dict)
            df_grouped_final['clean_netto'] = df_grouped_final['clean_bersih'] - df_grouped_final['clean_admin']

            display_cols = {}
            if col_notx: display_cols[col_notx] = "No. Transaksi"
            if col_tgl: display_cols[col_tgl] = "Tanggal"
            if col_pasien: display_cols[col_pasien] = "Nama Pasien"
            if col_jenis: display_cols[col_jenis] = "Cara Bayar"
            display_cols['clean_bersih'] = "Nominal (Rp)"
            display_cols['clean_admin'] = "Admin EDC/QRIS (Rp)"
            display_cols['clean_netto'] = "Netto Setelah Potongan (Rp)"

            df_display_clean = df_grouped_final[list(display_cols.keys())].rename(columns=display_cols)

            if col_jenis and col_jenis in df_grouped_final.columns:
                cash_mask = df_grouped_final[col_jenis].astype(str).str.upper().str.contains('CASH|TUNAI')
                penerimaan_tunai = float(df_grouped_final[cash_mask]['clean_bersih'].sum())
                
                nontunai_mask = ~cash_mask
                total_non_tunai_bruto = float(df_grouped_final[nontunai_mask]['clean_bersih'].sum())
                total_biaya_admin = float(df_grouped_final['clean_admin'].sum())
                
                st.sidebar.success(f"📊 Auto-rekap berhasil: {len(df_grouped_final)} transaksi unik terbaca.")
        else:
            st.sidebar.warning("⚠️ Data transaksi valid tidak ditemukan dalam file SIMRS.")

    with col2:
        st.subheader("💰 Transaksi Tunai (Rp)")
        modal_awal = st.number_input("Saldo Awal Kas Shift (Modal)", value=1141700.0, step=50000.0)
        penerimaan_tunai = st.number_input("Penerimaan Tunai Pelayanan", value=penerimaan_tunai, step=10000.0)
        piutang_tunai = st.number_input("Pelunasan Piutang Tunai", value=0.0, step=10000.0)
        deposit_tunai = st.number_input("Penerimaan Deposit Tunai", value=0.0, step=10000.0)
        refund_tunai = st.number_input("Dikurangi: Refund Tunai", value=0.0, step=10000.0)
        
        total_tunai_netto = penerimaan_tunai + piutang_tunai + deposit_tunai - refund_tunai
        st.success(f" Total Netto Tunai: **Rp {total_tunai_netto:,.2f}**")

    with col3:
        st.subheader("💳 Transaksi Non-Tunai (Rp)")
        penerimaan_nontunai_pelayanan = st.number_input("Penerimaan Non-Tunai Pelayanan", value=total_non_tunai_bruto, step=50000.0)
        piutang_nontunai = st.number_input("Pelunasan Piutang Non-Tunai", value=0.0, step=10000.0)
        deposit_nontunai = st.number_input("Penerimaan Deposit Non-Tunai", value=0.0, step=10000.0)
        refund_nontunai = st.number_input("Dikurangi: Refund Non-Tunai", value=0.0, step=10000.0)
        total_biaya_admin = st.number_input("Potongan Biaya Admin EDC/QRIS", value=total_biaya_admin, step=1000.0)
        
        total_nontunai_bruto_all = penerimaan_nontunai_pelayanan + piutang_nontunai + deposit_nontunai - refund_nontunai
        total_non_tunai_netto = total_nontunai_bruto_all - total_biaya_admin
        st.info(f" Total Netto Non-Tunai: **Rp {total_non_tunai_netto:,.2f}**")

# ==========================================
# MODUL RINGKASAN DATA OLAHAN SIMRS
# ==========================================
if file_parsed and not df_display_clean.empty:
    with st.expander("🔍 Ringkasan Data Transaksi SIMRS (Fokus: Nominal, Admin, & Netto)", expanded=True):
        st.dataframe(df_display_clean, use_container_width=True)
        csv_olahan = df_display_clean.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Unduh Ringkasan SIMRS (CSV)",
            data=csv_olahan,
            file_name=f"Ringkasan_SIMRS_{tgl_shift}.csv",
            mime="text/csv"
        )

# Cash Breakdown & Calculations
st.markdown("---")
st.subheader("💵 Input Pecahan Uang Fisik Kasir (Cash Count)")
col_pec1, col_pec2, col_pec3, col_pec4 = st.columns(4)

with col_pec1:
    l100 = st.number_input("100.000 (Lembar)", min_value=0, value=8)
    l50 = st.number_input("50.000 (Lembar)", min_value=0, value=6)
with col_pec2:
    l20 = st.number_input("20.000 (Lembar)", min_value=0, value=2)
    l10 = st.number_input("10.000 (Lembar)", min_value=0, value=13)
with col_pec3:
    l5 = st.number_input("5.000 (Lembar)", min_value=0, value=16)
    l2 = st.number_input("2.000 (Lembar)", min_value=0, value=18)
with col_pec4:
    l1 = st.number_input("1.000 (Lembar/Keping)", min_value=0, value=1)
    logam = st.number_input("Total Uang Logam (Rp)", min_value=0.0, value=51700.0, step=100.0)

total_kas_seharusnya = modal_awal + total_tunai_netto
total_uang_fisik = (l100 * 100000) + (l50 * 50000) + (l20 * 20000) + (l10 * 10000) + (l5 * 5000) + (l2 * 2000) + (l1 * 1000) + logam
selisih_kas = total_uang_fisik - total_kas_seharusnya
total_pendapatan_netto = total_tunai_netto + total_non_tunai_netto

st.markdown("---")
st.subheader("⏳ Transaksi / Tagihan Dalam Proses (Pending / Outstanding)")
initial_pending_data = pd.DataFrame([
    {"No. Transaksi / RM": "TRX-00129", "Keterangan / Kendala": "Menunggu konfirmasi settlement EDC", "Nominal (Rp)": 250000.0}
])
edited_pending_df = st.data_editor(
    initial_pending_data,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Nominal (Rp)": st.column_config.NumberColumn(format="Rp %.2f", step=10000.0)
    }
)
total_nominal_pending = float(edited_pending_df["Nominal (Rp)"].sum()) if not edited_pending_df.empty else 0.0

st.markdown("---")
st.subheader("📊 Hasil Rekonsiliasi Otomatis Shift")
m_col1, m_col2, m_col3, m_col4 = st.columns(4)
with m_col1:
    st.metric(label="Total Kas Fisik (Modal+Tunai)", value=f"Rp {total_kas_seharusnya:,.2f}")
with m_col2:
    st.metric(label="Uang Tunai Aktual (Brankas)", value=f"Rp {total_uang_fisik:,.2f}")
with m_col3:
    status_selisih = "PAS / SESUAI" if selisih_kas == 0 else ("LEBIH" if selisih_kas > 0 else "KURANG")
    st.metric(label=f"Selisih Kas ({status_selisih})", value=f"Rp {selisih_kas:,.2f}")
with m_col4:
    st.metric(label="Total Admin EDC / QRIS", value=f"Rp {total_biaya_admin:,.2f}")

st.markdown("---")
st.subheader("📝 Catatan Tambahan Kasir")
catatan_tambahan = st.text_area("Catatan Tambahan", "Uang lebih Rp 22 karena pasien tidak mau menerima kembalian")

# Generate PDF function
def create_pdf():
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=14, alignment=1, spaceAfter=2)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, alignment=1, spaceAfter=10)
    normal_bold = ParagraphStyle('NormalBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9)
    normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=8)
    
    elements = []
    elements.append(Paragraph("RUMAH SAKIT ADHYAKSA JAWA TIMUR", title_style))
    elements.append(Paragraph("FORMULIR SERAH TERIMA & CLOSING KASIR SHIFT", subtitle_style))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1e4d2b"), spaceAfter=10))
    
    meta_data = [
        [Paragraph("<b>Tanggal Shift:</b>", normal_style), Paragraph(str(tgl_shift), normal_style), Paragraph("<b>Petugas Menyerahkan:</b>", normal_style), Paragraph(petugas_lama, normal_style)],
        [Paragraph("<b>Shift Operasional:</b>", normal_style), Paragraph(shift_opt, normal_style), Paragraph("<b>Petugas Menerima:</b>", normal_style), Paragraph(petugas_baru, normal_style)]
    ]
    t_meta = Table(meta_data, colWidths=[100, 170, 110, 170])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F2F4F3")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_meta)
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b>1. REKAPITULASI PENERIMAAN KAS & NON-TUNAI</b>", normal_bold))
    rekap_data = [
        ["Uraian Penerimaan", "Penerimaan Tunai (Rp)", "Penerimaan Non-Tunai (Rp)", "Total Netto (Rp)"],
        ["Saldo Awal Kas / Modal Kembalian", f"{modal_awal:,.2f}", "-", f"{modal_awal:,.2f}"],
        ["1. Penerimaan Pelayanan", f"{penerimaan_tunai:,.2f}", f"{penerimaan_nontunai_pelayanan:,.2f}", f"{(penerimaan_tunai + penerimaan_nontunai_pelayanan):,.2f}"],
        ["2. Pelunasan Piutang", f"{piutang_tunai:,.2f}", f"{piutang_nontunai:,.2f}", f"{(piutang_tunai + piutang_nontunai):,.2f}"],
        ["3. Penerimaan Deposit", f"{deposit_tunai:,.2f}", f"{deposit_nontunai:,.2f}", f"{(deposit_tunai + deposit_nontunai):,.2f}"],
        ["4. Dikurangi: Batal / Refund", f"({refund_tunai:,.2f})", f"({refund_nontunai:,.2f})", f"({(refund_tunai + refund_nontunai):,.2f})"],
        ["5. Dikurangi: Potongan Admin EDC / QRIS", "-", f"({total_biaya_admin:,.2f})", f"({total_biaya_admin:,.2f})"],
        ["GRAND TOTAL PENDAPATAN SHIFT", f"{total_tunai_netto:,.2f}", f"{total_non_tunai_netto:,.2f}", f"{total_pendapatan_netto:,.2f}"]
    ]
    t_rekap = Table(rekap_data, colWidths=[210, 110, 115, 115])
    t_rekap.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e4d2b")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#E8F5E9")),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_rekap)
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b>2. RINCIAN UANG FISIK BRANKAS (CASH COUNT)</b>", normal_bold))
    cash_data = [
        ["Pecahan", "Jumlah", "Total Nominal", "Pecahan", "Jumlah", "Total Nominal"],
        ["Rp 100.000", str(l100), f"Rp {l100*100000:,.2f}", "Rp 5.000", str(l5), f"Rp {l5*5000:,.2f}"],
        ["Rp 50.000", str(l50), f"Rp {l50*50000:,.2f}", "Rp 2.000", str(l2), f"Rp {l2*2000:,.2f}"],
        ["Rp 20.000", str(l20), f"Rp {l20*20000:,.2f}", "Rp 1000", str(l1), f"Rp {l1*1000:,.2f}"],
        ["Rp 10.000", str(l10), f"Rp {l10*10000:,.2f}", "Uang Logam", "-", f"Rp {logam:,.2f}"],
        ["TOTAL UANG FISIK AKTUAL", "", "", "", "", f"Rp {total_uang_fisik:,.2f}"],
        ["KAS SEHARUSNYA (MODAL + TUNAI)", "", "", "", "", f"Rp {total_kas_seharusnya:,.2f}"],
        [f"SELISIH KAS ({status_selisih})", "", "", "", "", f"Rp {selisih_kas:,.2f}"]
    ]
    t_cash = Table(cash_data, colWidths=[90, 50, 135, 90, 50, 135])
    t_cash.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#444444")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,0), (2,-1), 'RIGHT'),
        ('ALIGN', (4,0), (5,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('SPAN', (0,5), (4,5)),
        ('SPAN', (0,6), (4,6)),
        ('SPAN', (0,7), (4,7)),
        ('FONTNAME', (0,5), (-1,-1), 'Helvetica-Bold'),
        ('BACKGROUND', (0,7), (-1,7), colors.HexColor("#FFF3E0")),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_cash)
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("<b>3. TRANSAKSI / TAGIHAN DALAM PROSES (PENDING / OUTSTANDING)</b>", normal_bold))
    pending_table_data = [["No.", "No. Transaksi / RM", "Keterangan / Kendala", "Nominal (Rp)"]]
    
    if not edited_pending_df.empty:
        for idx, row in edited_pending_df.reset_index(drop=True).iterrows():
            nom = row.get("Nominal (Rp)", 0.0)
            pending_table_data.append([
                str(idx + 1),
                str(row.get("No. Transaksi / RM", "")),
                str(row.get("Keterangan / Kendala", "")),
                f"Rp {nom:,.2f}"
            ])
        pending_table_data.append(["", "", "TOTAL NOMINAL DALAM PROSES", f"Rp {total_nominal_pending:,.2f}"])
    else:
        pending_table_data.append(["-", "Tidak ada transaksi pending", "-", "Rp 0.00"])

    t_pending = Table(pending_table_data, colWidths=[30, 130, 230, 114])
    t_pending.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e4d2b")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (3,1), (3,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#E3F2FD")),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_pending)
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph(f"<b>Catatan Kasir:</b> {catatan_tambahan}", normal_style))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b>4. PERNYATAAN SERAH TERIMA ANTAR SHIFT</b>", normal_bold))
    pernyataan_text = "Kas, dokumen, dan informasi transaksi shift telah diperiksa dan diserahterimakan sesuai kondisi pada saat pergantian shift."
    elements.append(Paragraph(pernyataan_text, normal_style))
    elements.append(Spacer(1, 15))
    
    pj_name = pj_kasir if pj_kasir.strip() != "" else " ( .................................... ) "
    sig_data = [
        ["Petugas Shift Lama (Menyerahkan)", "Petugas Shift Baru (Menerima)", "Mengetahui (Penanggung Jawab Kasir)"],
        ["\n\n\n________________________", "\n\n\n________________________", "\n\n\n________________________"],
        [petugas_lama, petugas_baru, pj_name]
    ]
    t_sig = Table(sig_data, colWidths=[180, 180, 190])
    t_sig.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),
    ]))
    elements.append(t_sig)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

st.markdown("---")
st.subheader("🖨️ Cetak & Unduh Dokumen Closing")
pdf_bytes = create_pdf()

st.download_button(
    label="📄 Unduh Form Closing Kasir (PDF)",
    data=pdf_bytes,
    file_name=f"Serah_Terima_Kasir_{tgl_shift}.pdf",
    mime="applicatio
