import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(
    page_title="Rekonsiliasi Cashless — ASDP Ternate",
    page_icon="🚢",
    layout="wide",
)

# ── CSS ───────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: #2563EB;
        color: white;
        padding: 16px 24px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    .main-header h1 { margin: 0; font-size: 22px; }
    .main-header p  { margin: 4px 0 0; font-size: 13px; opacity: .75; }
    .metric-card {
        background: #F3F4F6;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 14px 16px;
    }
    .metric-label { font-size: 13px; color: #6B7280; margin-bottom: 4px; }
    .metric-value { font-size: 20px; font-weight: 700; }
    .badge-match    { background:#F0FDF4; color:#16A34A; padding:3px 10px; border-radius:12px; font-size:12px; font-weight:600; }
    .badge-no-set   { background:#FEF2F2; color:#DC2626; padding:3px 10px; border-radius:12px; font-size:12px; font-weight:600; }
    .badge-no-tiket { background:#FFFBEB; color:#D97706; padding:3px 10px; border-radius:12px; font-size:12px; font-weight:600; }
    .badge-prepaid  { background:#F5F3FF; color:#7C3AED; padding:3px 10px; border-radius:12px; font-size:12px; font-weight:600; }
    .info-box {
        background: #EFF6FF;
        border: 1px solid #93C5FD;
        border-radius: 8px;
        padding: 14px 18px;
        color: #1E40AF;
        font-size: 14px;
    }
    .bank-card {
        background: #EFF6FF;
        border: 1px solid #93C5FD;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
    div[data-testid="stTabs"] button { font-size: 14px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────
now = datetime.now().strftime("%A, %d %B %Y  %H:%M")
st.markdown(f"""
<div class="main-header">
  <h1>🚢 Rekonsiliasi Cashless ASDP — Cabang Ternate</h1>
  <p>{now}</p>
</div>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────
for k in ("df_rekon", "df_settle_raw", "df_prepaid", "processed"):
    if k not in st.session_state:
        st.session_state[k] = None
if "processed" not in st.session_state:
    st.session_state.processed = False

# ── Tabs ──────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["   📂 Upload File   ", "   🔍 Rekonsiliasi Order ID   ", "   📊 Ringkasan   "])

# ══════════════════════════════════════════════════════════════
# TAB 1 — Upload File
# ══════════════════════════════════════════════════════════════
with tab1:
    st.markdown("### Upload File Excel")
    st.markdown("Pilih file Excel untuk **Tiket Detail** dan **Settlement Espay**. Boleh file yang sama.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**📄 Tiket Detail**")
        st.caption("Sheet: `Tiket Detail` — Kolom: Order ID, Golongan, Nominal, St Kapal, St Bayar")
        file_tiket = st.file_uploader("Upload file Tiket Detail", type=["xlsx", "xls"], key="tiket")

    with col2:
        st.markdown("**📄 Settlement Espay**")
        st.caption("Sheet: `Settlement Espay` — Kolom: Transaction ID, Order Id, Amount, Bank, Settlement Date")
        file_settle = st.file_uploader("Upload file Settlement Espay", type=["xlsx", "xls"], key="settle")

    st.markdown("---")

    st.markdown("""
<div class="info-box">
<b>Cara Pakai:</b><br>
1. Upload file Excel Rekonsiliasi Cabang (untuk Tiket Detail).<br>
2. Upload file yang sama atau file lain untuk Settlement Espay.<br>
3. Klik <b>Proses Rekonsiliasi</b> — hasil muncul di tab Rekonsiliasi Order ID dan Ringkasan.
</div>
""", unsafe_allow_html=True)

    st.markdown("")
    btn_proses = st.button("▶️  Proses Rekonsiliasi", type="primary", use_container_width=False)

    if btn_proses:
        if not file_tiket or not file_settle:
            st.error("⚠️ Pilih file untuk Tiket Detail dan Settlement Espay terlebih dahulu.")
        else:
            with st.spinner("Membaca dan memproses data..."):
                try:
                    # Tiket Detail
                    df_t = pd.read_excel(file_tiket, sheet_name="Tiket Detail", dtype={"Order ID": str})
                    nom_col = df_t.columns[24]
                    df_t = df_t.rename(columns={nom_col: "Nominal"})
                    df_t = df_t[df_t["St Bayar"].astype(str).str.lower().str.strip() == "paid"].copy()
                    df_t["Order ID"] = df_t["Order ID"].astype(str).str.strip()
                    df_t["Bank"] = df_t["Bank"].astype(str).str.strip()

                    # Pisahkan ESPAY vs Prepaid
                    df_espay   = df_t[df_t["Bank"].str.upper() == "ESPAY"].copy()
                    df_prepaid_raw = df_t[df_t["Bank"].str.upper() != "ESPAY"].copy()

                    # Ringkasan prepaid per Bank
                    prepaid_grp = (df_prepaid_raw.groupby("Bank")
                                   .agg(jumlah_tiket=("Nominal", "count"),
                                        nominal=("Nominal", "sum"))
                                   .reset_index()
                                   .rename(columns={"jumlah_tiket": "Jml Tiket", "nominal": "Nominal (Rp)"}))
                    st.session_state.df_prepaid = prepaid_grp

                    # Settlement Espay
                    df_s = pd.read_excel(file_settle, sheet_name="Settlement Espay", dtype={"Order Id": str})
                    tipe_col = df_s.columns[24]
                    df_s = df_s.rename(columns={tipe_col: "Tipe", "Order Id": "Order ID"})
                    df_s["Order ID"] = df_s["Order ID"].astype(str).str.strip()
                    st.session_state.df_settle_raw = df_s.copy()

                    # Grouping — hanya dari tiket ESPAY
                    tiket_grp = (df_espay.groupby("Order ID")
                                 .agg(jumlah_tiket=("Nominal", "count"),
                                      nominal_tiket=("Nominal", "sum"))
                                 .reset_index())

                    settle_grp = (df_s.groupby("Order ID")
                                  .agg(jumlah_settle=("Amount", "count"),
                                       nominal_settle=("Amount", "sum"),
                                       bank=("BANK", lambda x: x.mode().iloc[0] if len(x) else ""),
                                       tipe=("Tipe", lambda x: x.mode().iloc[0] if len(x) else ""))
                                  .reset_index())

                    df_m = pd.merge(tiket_grp, settle_grp, on="Order ID", how="outer")
                    df_m["jumlah_tiket"]   = df_m["jumlah_tiket"].fillna(0).astype(int)
                    df_m["nominal_tiket"]  = df_m["nominal_tiket"].fillna(0)
                    df_m["jumlah_settle"]  = df_m["jumlah_settle"].fillna(0).astype(int)
                    df_m["nominal_settle"] = df_m["nominal_settle"].fillna(0)
                    df_m["bank"]           = df_m["bank"].fillna("")
                    df_m["tipe"]           = df_m["tipe"].fillna("")
                    df_m["selisih"]        = df_m["nominal_tiket"] - df_m["nominal_settle"]
                    df_m["_ada_tiket"]     = df_m["jumlah_tiket"] > 0
                    df_m["_ada_settle"]    = df_m["jumlah_settle"] > 0

                    def get_status(r):
                        if r["_ada_tiket"] and r["_ada_settle"]:
                            return "Match"
                        elif r["_ada_tiket"]:
                            return "Tidak di Settlement"
                        else:
                            return "Tidak di Tiket Detail"

                    df_m["Status"] = df_m.apply(get_status, axis=1)
                    df_m = df_m.rename(columns={
                        "jumlah_tiket":   "Jml Tiket",
                        "nominal_tiket":  "Nominal Tiket (Rp)",
                        "jumlah_settle":  "Jml Settlement",
                        "nominal_settle": "Nominal Settlement (Rp)",
                        "selisih":        "Selisih (Rp)",
                        "bank":           "Bank",
                        "tipe":           "Tipe",
                    }).reset_index(drop=True)

                    st.session_state.df_rekon = df_m
                    st.session_state.processed = True

                    match_c    = (df_m["Status"] == "Match").sum()
                    no_set_c   = (df_m["Status"] == "Tidak di Settlement").sum()
                    no_tkt_c   = (df_m["Status"] == "Tidak di Tiket Detail").sum()
                    prepaid_c  = df_prepaid_raw["Order ID"].nunique()

                    st.success(
                        f"✅ Selesai — {len(df_m):,} unique Order ID ESPAY  |  "
                        f"Match: {match_c:,}  |  "
                        f"Tidak di Settlement: {no_set_c:,}  |  "
                        f"Tidak di Tiket Detail: {no_tkt_c:,}  |  "
                        f"Prepaid (non-ESPAY): {prepaid_c:,} Order ID"
                    )
                    st.info("Buka tab **Rekonsiliasi Order ID** atau **Ringkasan** untuk melihat hasil.")

                except Exception as e:
                    import traceback
                    st.error(f"❌ Error saat memproses file:\n\n```\n{traceback.format_exc()}\n```")

# ══════════════════════════════════════════════════════════════
# TAB 2 — Rekonsiliasi Order ID
# ══════════════════════════════════════════════════════════════
with tab2:
    if st.session_state.df_rekon is None:
        st.info("Belum ada data. Silakan upload file dan proses rekonsiliasi di tab **Upload File**.")
    else:
        df = st.session_state.df_rekon.copy()

        # Filter & Search
        col_f, col_s, col_cnt = st.columns([2, 3, 2])
        with col_f:
            filter_opt = st.selectbox(
                "Filter Status",
                ["Semua", "Match", "Tidak di Settlement", "Tidak di Tiket Detail"],
                label_visibility="collapsed"
            )
        with col_s:
            search_q = st.text_input("Cari Order ID", placeholder="🔍 Cari Order ID...", label_visibility="collapsed")

        # Apply filter
        if filter_opt != "Semua":
            df = df[df["Status"] == filter_opt]
        if search_q.strip():
            df = df[df["Order ID"].str.lower().str.contains(search_q.strip().lower(), na=False)]

        with col_cnt:
            st.markdown(f"**{len(df):,}** Order ID ditampilkan")

        # Legend
        st.markdown(
            '<span class="badge-match">● Match</span> &nbsp;'
            '<span class="badge-no-set">● Tidak di Settlement</span> &nbsp;'
            '<span class="badge-no-tiket">● Tidak di Tiket Detail</span>',
            unsafe_allow_html=True
        )
        st.markdown("")

        # Prepare display dataframe
        disp = df[["Order ID", "Status", "Jml Tiket", "Nominal Tiket (Rp)",
                   "Jml Settlement", "Nominal Settlement (Rp)", "Selisih (Rp)", "Bank", "Tipe"]].copy()
        disp.insert(0, "No", range(1, len(disp) + 1))

        def color_status(val):
            if val == "Match":
                return "background-color: #F0FDF4; color: #16A34A"
            elif val == "Tidak di Settlement":
                return "background-color: #FEF2F2; color: #DC2626"
            elif val == "Tidak di Tiket Detail":
                return "background-color: #FFFBEB; color: #D97706"
            return ""

        styled = disp.style.map(color_status, subset=["Status"]).format({
            "Nominal Tiket (Rp)": "{:,.0f}",
            "Nominal Settlement (Rp)": "{:,.0f}",
            "Selisih (Rp)": "{:,.0f}",
        })

        st.dataframe(styled, use_container_width=True, height=520, hide_index=True)

        # Download
        st.markdown("---")
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df.drop(columns=["_ada_tiket", "_ada_settle"], errors="ignore").to_excel(w, index=False, sheet_name="Rekonsiliasi")
            st.download_button(
                "⬇️ Download Hasil (Excel)",
                data=buf.getvalue(),
                file_name=f"rekonsiliasi_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_dl2:
            csv_buf = df.drop(columns=["_ada_tiket", "_ada_settle"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Download Hasil (CSV)",
                data=csv_buf,
                file_name=f"rekonsiliasi_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

# ══════════════════════════════════════════════════════════════
# TAB 3 — Ringkasan
# ══════════════════════════════════════════════════════════════
with tab3:
    if st.session_state.df_rekon is None:
        st.info("Belum ada data. Proses rekonsiliasi terlebih dahulu di tab **Upload File**.")
    else:
        df      = st.session_state.df_rekon
        df_s_raw = st.session_state.df_settle_raw

        match_c    = (df["Status"] == "Match").sum()
        no_set_c   = (df["Status"] == "Tidak di Settlement").sum()
        no_tkt_c   = (df["Status"] == "Tidak di Tiket Detail").sum()
        nom_tiket  = df["Nominal Tiket (Rp)"].sum()
        nom_settle = df["Nominal Settlement (Rp)"].sum()
        selisih    = nom_tiket - nom_settle
        nom_hilang = df[df["Status"] == "Tidak di Settlement"]["Nominal Tiket (Rp)"].sum()
        nom_lebih  = df[df["Status"] == "Tidak di Tiket Detail"]["Nominal Settlement (Rp)"].sum()
        total_uord = df[df["_ada_tiket"]]["Order ID"].nunique()
        total_sord = df[df["_ada_settle"]]["Order ID"].nunique()

        st.markdown("### Ringkasan Rekonsiliasi")

        # Row 1 — jumlah order
        c1, c2, c3, c4, c5 = st.columns(5)
        metrics_row1 = [
            (c1, "Unique Order — Tiket Detail", f"{total_uord:,}", "#2563EB"),
            (c2, "Unique Order — Settlement",   f"{total_sord:,}", "#2563EB"),
            (c3, "Order Match",                 f"{match_c:,}",   "#16A34A"),
            (c4, "Hilang di Settlement",        f"{no_set_c:,}",  "#DC2626"),
            (c5, "Lebih di Settlement",         f"{no_tkt_c:,}",  "#D97706"),
        ]
        for col, label, value, color in metrics_row1:
            with col:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value" style="color:{color}">{value}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("")

        # Row 2 — nominal
        d1, d2, d3, d4, d5 = st.columns(5)
        sel_color = "#DC2626" if selisih != 0 else "#16A34A"
        metrics_row2 = [
            (d1, "Total Nominal Tiket",          f"Rp {nom_tiket:,.0f}",  "#111827"),
            (d2, "Total Nominal Settlement",     f"Rp {nom_settle:,.0f}", "#111827"),
            (d3, "Selisih (Tiket - Settlement)", f"Rp {selisih:,.0f}",    sel_color),
            (d4, "Nominal Order Hilang",         f"Rp {nom_hilang:,.0f}", "#DC2626"),
            (d5, "Nominal Order Lebih",          f"Rp {nom_lebih:,.0f}",  "#D97706"),
        ]
        for col, label, value, color in metrics_row2:
            with col:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value" style="color:{color}; font-size:15px">{value}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("---")

        # Breakdown Bank
        st.markdown("#### Breakdown per Bank (Settlement)")
        bank_grp = (df_s_raw.groupby("BANK")["Amount"]
                    .agg(["sum", "count"])
                    .reset_index()
                    .sort_values("sum", ascending=False))
        bank_cols = st.columns(min(len(bank_grp), 5))
        for i, (_, row) in enumerate(bank_grp.iterrows()):
            if i >= len(bank_cols):
                break
            with bank_cols[i]:
                st.markdown(f"""
                <div class="bank-card">
                    <div style="font-size:13px;color:#6B7280">{row['BANK']}</div>
                    <div style="font-size:16px;font-weight:700;color:#2563EB">Rp {row['sum']:,.0f}</div>
                    <div style="font-size:12px;color:#6B7280">{int(row['count']):,} transaksi</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("")

        # Breakdown Tipe
        st.markdown("#### Breakdown Tipe Transaksi (Settlement)")
        tipe_grp = (df_s_raw.groupby("Tipe")["Amount"]
                    .agg(["sum", "count"])
                    .reset_index()
                    .sort_values("sum", ascending=False))
        tipe_colors = {"Go Show": ("#FFFBEB", "#D97706"), "Online": ("#EFF6FF", "#2563EB")}
        tipe_cols = st.columns(min(len(tipe_grp), 5))
        for i, (_, row) in enumerate(tipe_grp.iterrows()):
            if i >= len(tipe_cols):
                break
            bg_c, fg_c = tipe_colors.get(str(row["Tipe"]), ("#F3F4F6", "#6B7280"))
            with tipe_cols[i]:
                st.markdown(f"""
                <div style="background:{bg_c};border:1px solid #E5E7EB;border-radius:8px;padding:12px 16px;margin-bottom:8px">
                    <div style="font-size:13px;color:#6B7280">{row['Tipe']}</div>
                    <div style="font-size:16px;font-weight:700;color:{fg_c}">Rp {row['sum']:,.0f}</div>
                    <div style="font-size:12px;color:#6B7280">{int(row['count']):,} transaksi</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("---")

        # ── Prepaid (non-ESPAY) ───────────────────────────────
        df_prepaid = st.session_state.df_prepaid
        if df_prepaid is not None and len(df_prepaid):
            st.markdown("#### 🟣 Prepaid (Transaksi Non-ESPAY di Tiket Detail)")

            total_prepaid_tiket = df_prepaid["Jml Tiket"].sum()
            total_prepaid_nom   = df_prepaid["Nominal (Rp)"].sum()

            p1, p2, p3 = st.columns(3)
            with p1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Jumlah Bank Prepaid</div>
                    <div class="metric-value" style="color:#7C3AED">{len(df_prepaid):,}</div>
                </div>""", unsafe_allow_html=True)
            with p2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Total Tiket Prepaid</div>
                    <div class="metric-value" style="color:#7C3AED">{total_prepaid_tiket:,}</div>
                </div>""", unsafe_allow_html=True)
            with p3:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Total Nominal Prepaid</div>
                    <div class="metric-value" style="color:#7C3AED; font-size:15px">Rp {total_prepaid_nom:,.0f}</div>
                </div>""", unsafe_allow_html=True)

            st.markdown("")
            st.dataframe(
                df_prepaid.style.format({"Nominal (Rp)": "{:,.0f}"}),
                use_container_width=True,
                hide_index=True,
            )
            st.markdown("---")

        # Tabel order hilang
        hilang_df = df[df["Status"] == "Tidak di Settlement"][
            ["Order ID", "Nominal Tiket (Rp)", "Jml Tiket"]].head(200)
        if len(hilang_df):
            st.markdown(f"#### 🔴 Order ID Tidak Ada di Settlement ({no_set_c:,} total, menampilkan maks 200)")
            st.dataframe(
                hilang_df.style.format({"Nominal Tiket (Rp)": "{:,.0f}"}),
                use_container_width=True,
                hide_index=True,
                height=300,
            )
