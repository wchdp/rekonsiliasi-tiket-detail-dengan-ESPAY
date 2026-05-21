import streamlit as st
import pandas as pd
import io
import re
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
    .warn-box {
        background: #FFF7ED;
        border: 1px solid #FDBA74;
        border-radius: 8px;
        padding: 14px 18px;
        color: #9A3412;
        font-size: 14px;
    }
    .bank-card {
        background: #EFF6FF;
        border: 1px solid #93C5FD;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
    .green-card {
        background: #F0FDF4;
        border: 1px solid #86EFAC;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
    .purple-card {
        background: #F5F3FF;
        border: 1px solid #C4B5FD;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
    .red-card {
        background: #FEF2F2;
        border: 1px solid #FECACA;
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
for k in (
    "df_rekon", "df_settle_raw", "df_prepaid", "df_prepaid_raw",
    "df_espay", "processed", "periode_min", "periode_max",
    "settle_total", "settle_filtered",
    "df_bank_bri", "df_bank_bca", "df_bank_mandiri",
    "bank_processed",
):
    if k not in st.session_state:
        st.session_state[k] = None
if "processed" not in st.session_state:
    st.session_state.processed = False
if "bank_processed" not in st.session_state:
    st.session_state.bank_processed = False

# ══════════════════════════════════════════════════════════════
# HELPER — Klasifikasi baris rekening koran
# ══════════════════════════════════════════════════════════════

def extract_mrc_id(keterangan: str) -> str | None:
    """Ekstrak 18 karakter mulai dari 'mrc' di kolom keterangan."""
    kl = str(keterangan).lower()
    idx = kl.find("mrc")
    if idx == -1:
        return None
    return keterangan[idx: idx + 18]

def classify_bank_row(keterangan: str, nominal: float, bank_type: str,
                      espay_settlement_df: pd.DataFrame | None) -> dict:
    """
    Klasifikasi satu baris rekening koran sesuai rumus Excel di sheet BCA dan BRI CASHLESS.

    Rumus Excel (col K = Kategori, col L = Tipe) di kedua sheet identik polanya:

    BRI (col K):
      =IF(ISNUMBER(SEARCH("mrc",$D)),
          IFERROR(INDEX(Settlement!$AA, MATCH(MID($D,SEARCH("mrc",$D),18), Settlement!$R, 0)),
                  "Prepaid"),
          IF(ISNUMBER(SEARCH("Pinbuk",$D)), "NON BCA", "Prepaid"))

    BCA (col K):
      =IF(ISNUMBER(SEARCH("mrc",$D)),
          IFERROR(INDEX(Settlement!$AA, MATCH(MID($D,SEARCH("mrc",$D),18), Settlement!$R, 0)),
                  "Cash"),
          IF(AND(ISNUMBER(SEARCH("KR OTOMATIS",$D)), ISNUMBER(SEARCH("MID",$D))), "BCA", "Cash"))

    Settlement Espay:
      - Col R (idx 17) = Settlement Remark  → berisi nilai mrc... untuk di-MATCH
      - Col Y (idx 24) = Type               → "Go Show" / "Online"
      - Col AA (idx 26) = BANK              → "BCA", "NON BCA", dll

    Jika ditemukan di Settlement:
      - kategori = nilai col AA (BANK), mis. "NON BCA", "BCA"
      - tipe     = nilai col Y  (Type), mis. "Go Show", "Online"
    Jika mrc ada tapi tidak ditemukan di Settlement:
      - BRI/MANDIRI → kategori = "Prepaid", tipe = "Prepaid"
      - BCA         → kategori = "Cash",    tipe = "Cash"
    Jika tidak ada mrc:
      - BRI: "Pinbuk" → kategori="NON BCA", tipe="Go Show" | lainnya → "Prepaid"
      - BCA: "KR OTOMATIS"+"MID" → kategori="BCA", tipe="Go Show" | lainnya → "Cash"
      - MANDIRI: "SWITCHING"/"EDC"/"MTRANSFER" → kategori="NON MANDIRI" | lainnya → "Prepaid"
    """
    ket = str(keterangan)
    mrc_id = extract_mrc_id(ket)

    kategori    = "Lainnya"
    tipe        = "-"
    order_match = None

    # Kolom Settlement Espay sesuai header Excel:
    # R = Settlement Remark (idx 17), Y = Type (idx 24), AA = BANK (idx 26)
    COL_R  = "Settlement Remark"   # col R, idx 17
    COL_Y  = "Type"                # col Y, idx 24 — nama duplikat; pakai posisi
    COL_AA = "BANK"                # col AA, idx 26

    if mrc_id and espay_settlement_df is not None and len(espay_settlement_df):
        cols = espay_settlement_df.columns.tolist()

        # Pakai nama kolom jika ada, fallback ke indeks posisi
        ref_col    = COL_R  if COL_R  in cols else (cols[17] if len(cols) > 17 else None)
        # col Y (Type) — ada dua kolom "Type" (T dan Y); ambil indeks 24
        tipe_col   = cols[24] if len(cols) > 24 else None
        bank_col   = COL_AA if COL_AA in cols else (cols[26] if len(cols) > 26 else None)

        if ref_col is not None:
            # Match persis 18 karakter mrc ke col R
            matched = espay_settlement_df[
                espay_settlement_df[ref_col].astype(str).str.strip() == mrc_id.strip()
            ]
            if len(matched):
                # Ditemukan → ambil BANK (col AA) sebagai kategori, Type (col Y) sebagai tipe
                order_match = mrc_id
                kategori = str(matched.iloc[0][bank_col]).strip() if bank_col else "ESPAY"
                tipe     = str(matched.iloc[0][tipe_col]).strip() if tipe_col else "-"
            else:
                # Tidak ditemukan di Settlement
                if bank_type in ("BRI", "MANDIRI"):
                    kategori = "Prepaid"
                    tipe     = "Prepaid"
                else:  # BCA
                    kategori = "Cash"
                    tipe     = "Cash"
        else:
            if bank_type in ("BRI", "MANDIRI"):
                kategori, tipe = "Prepaid", "Prepaid"
            else:
                kategori, tipe = "Cash", "Cash"

    else:
        # Tidak ada 'mrc' dalam keterangan
        ket_lower = ket.lower()

        if bank_type == "BRI":
            # Excel: IF(ISNUMBER(SEARCH("Pinbuk",$D)), "NON BCA", "Prepaid")
            # SEARCH di Excel case-insensitive
            if "pinbuk" in ket_lower:
                kategori = "NON BCA"
                tipe     = "Go Show"
            else:
                kategori = "Prepaid"
                tipe     = "Prepaid"

        elif bank_type == "BCA":
            # Excel: IF(AND(ISNUMBER(SEARCH("KR OTOMATIS",$D)), ISNUMBER(SEARCH("MID",$D))), "BCA", "Cash")
            ket_up = ket.upper()
            if "KR OTOMATIS" in ket_up and "MID" in ket_up:
                kategori = "BCA"
                tipe     = "Go Show"
            else:
                kategori = "Cash"
                tipe     = "Cash"

        elif bank_type == "MANDIRI":
            ket_up = ket.upper()
            if any(x in ket_up for x in ["SWITCHING", "EDC", "MTRANSFER", "TRANSFER"]):
                kategori = "NON MANDIRI"
                tipe     = "Go Show"
            else:
                kategori = "Prepaid"
                tipe     = "Prepaid"

    return {
        "kategori":  kategori,
        "tipe":      tipe,
        "order_ref": order_match,
    }


def _parse_tanggal(series: pd.Series) -> pd.Series:
    """
    Parse kolom tanggal dengan benar untuk semua format bank.

    Format yang ada:
    - BCA/Mandiri: YYYY-MM-DD (ISO) atau YYYY-MM-DD HH:MM:SS → dayfirst=False
    - BRI        : DD-MM-YYYY (misal '01-02-2026')           → dayfirst=True

    Deteksi dilakukan dengan melihat panjang bagian pertama sebelum '-':
    - len 4 → YYYY-MM-DD → dayfirst=False
    - len 1-2 → DD-MM-YYYY → dayfirst=True
    """
    sample = series.dropna().astype(str).head(10)
    use_dayfirst = False
    for val in sample:
        # Hilangkan bagian waktu (HH:MM:SS) jika ada
        val_clean = str(val).split(' ')[0].strip()
        parts = val_clean.replace('/', '-').split('-')
        if len(parts) == 3:
            try:
                p0 = parts[0]
                if len(p0) == 4:        # YYYY-MM-DD → ISO, dayfirst=False
                    use_dayfirst = False
                    break
                elif len(p0) <= 2:      # DD-MM-YYYY → dayfirst=True
                    use_dayfirst = True
                    break
            except Exception:
                pass
    try:
        return pd.to_datetime(series, dayfirst=use_dayfirst, errors="coerce")
    except Exception:
        return pd.to_datetime(series, errors="coerce")


def process_bank_file(file, bank_type: str,
                      espay_settlement_df: pd.DataFrame | None) -> pd.DataFrame:
    """
    Baca rekening koran excel untuk BCA, BRI, dan Mandiri.

    Format yang didukung (deteksi otomatis):
    - BCA  : header=0, kolom DATE | TIME | REMARK | SALDO AWAL | DEBET | CREDIT | SALDO AKHIR
    - BRI  : header=0, kolom Unnamed:0(tanggal) | REMARK | SALDO AWAL | DEBET | CREDIT
             Tanggal format DD-MM-YYYY → wajib dayfirst=True
    - MANDIRI: header=0, kolom Tanggal | Remark | Debit | Credit

    Hanya baris dengan CREDIT > 0 (uang masuk) yang diproses.
    """
    xl = pd.ExcelFile(file)

    # ── Pilih sheet ───────────────────────────────────────────
    target_sheet = xl.sheet_names[0]
    bank_keywords = {"BRI": ["bri"], "BCA": ["bca"], "MANDIRI": ["mandiri"]}
    for s in xl.sheet_names:
        sl = s.lower()
        if any(k in sl for k in bank_keywords.get(bank_type, [])):
            target_sheet = s
            break
        if any(x in sl for x in ["koran", "mutasi", "rekening"]):
            target_sheet = s
            break

    # ── Baca file ─────────────────────────────────────────────
    # Coba berbagai posisi header; pilih yang punya kolom kredit
    df = None
    for header_row in [0, 1, 2, 3, 13]:
        try:
            df_try = pd.read_excel(file, sheet_name=target_sheet,
                                   header=header_row, dtype=str)
            df_try.columns = [
                (c.strip() if isinstance(c, str) else c)
                for c in df_try.columns
            ]
            cols_l = [str(c).lower() for c in df_try.columns]
            has_credit = any(x in c for c in cols_l
                             for x in ["credit", "kredit", "masuk", "nominal"])
            if has_credit:
                df = df_try
                break
        except Exception:
            continue

    if df is None:
        df = pd.read_excel(file, sheet_name=target_sheet, dtype=str)
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

    # ── Rename kolom tanpa nama (BRI: kolom tanggal = Unnamed: 0) ──
    # Jika kolom pertama adalah Unnamed dan kolom kedua adalah REMARK/keterangan,
    # maka kolom pertama pasti Tanggal.
    unnamed_first = str(df.columns[0]).startswith("Unnamed") or df.columns[0] != df.columns[0]
    if unnamed_first:
        # Cek apakah kolom pertama berisi data tanggal
        sample_vals = df.iloc[:5, 0].dropna().astype(str).tolist()
        looks_like_date = any(
            any(c.isdigit() for c in v) and ("-" in v or "/" in v)
            for v in sample_vals
        )
        if looks_like_date:
            cols = list(df.columns)
            cols[0] = "Tanggal"
            df.columns = cols

    # ── Deteksi kolom Tanggal ──────────────────────────────────
    tgl_col = None
    for c in df.columns:
        cl = str(c).lower()
        if any(x in cl for x in ["date", "tanggal", "tgl"]):
            tgl_col = c
            break
    if tgl_col is None:
        tgl_col = df.columns[0]

    # ── Deteksi kolom Keterangan ──────────────────────────────
    ket_col = None
    for c in df.columns:
        cl = str(c).lower()
        if any(x in cl for x in ["remark", "keterangan", "deskripsi", "description", "uraian"]):
            ket_col = c
            break
    if ket_col is None:
        ket_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

    # ── Deteksi kolom Kredit (uang masuk) ─────────────────────
    # Cari "credit"/"kredit"/"masuk" — eksplisit hindari "debet"/"debit"
    nom_col = None
    for c in df.columns:
        cl = str(c).lower()
        if any(x in cl for x in ["credit", "kredit", "masuk"]) and \
           "debet" not in cl and "debit" not in cl:
            nom_col = c
            break
    if nom_col is None:
        for c in df.columns:
            cl = str(c).lower()
            if any(x in cl for x in ["nominal", "amount", "jumlah"]) and \
               "debet" not in cl and "debit" not in cl:
                nom_col = c
                break
    if nom_col is None:
        # Fallback posisi: col G (index 6) untuk BCA/BRI, col D (index 3) untuk Mandiri
        nom_col = df.columns[min(6, len(df.columns) - 1)]

    # ── Susun dataframe kerja ──────────────────────────────────
    df = df[[tgl_col, ket_col, nom_col]].copy()
    df.columns = ["Tanggal", "Keterangan", "Nominal"]

    # ── Bersihkan nominal ──────────────────────────────────────
    df["Nominal"] = (
        df["Nominal"]
        .astype(str)
        .str.replace(r"[^\d\.]", "", regex=True)
        .replace("", "0")
    )
    df["Nominal"] = pd.to_numeric(df["Nominal"], errors="coerce").fillna(0)

    # ── Parse tanggal (handle DD-MM-YYYY untuk BRI) ───────────
    df["Tanggal"] = _parse_tanggal(df["Tanggal"])

    # ── Filter: uang masuk (Nominal > 0) dan tanggal valid ────
    df = df[(df["Nominal"] > 0) & (df["Tanggal"].notna())].copy()
    df = df.reset_index(drop=True)

    if len(df) == 0:
        return df

    # ── Klasifikasi ────────────────────────────────────────────
    results = df.apply(
        lambda row: classify_bank_row(
            row["Keterangan"], row["Nominal"], bank_type, espay_settlement_df
        ),
        axis=1,
        result_type="expand",
    )
    df["Kategori"]  = results["kategori"]
    df["Tipe"]      = results["tipe"]
    df["Order Ref"] = results["order_ref"]
    df["Bank"]      = bank_type

    return df


def detect_prepaid_shortfall(df_prepaid_raw: pd.DataFrame,
                              df_bank: pd.DataFrame | None) -> pd.DataFrame:
    """
    Bandingkan nominal kumulatif harian Prepaid di tiket detail
    dengan uang masuk Prepaid di rekening koran (kategori Prepaid / OnUs).
    Jika ada tanggal yang nominal bank < nominal tiket → kurang settlement.
    """
    if df_prepaid_raw is None or len(df_prepaid_raw) == 0:
        return pd.DataFrame()

    # Hitung kumulatif harian dari tiket detail prepaid
    df_t = df_prepaid_raw.copy()
    # Coba kolom tanggal transaksi
    date_col = None
    for c in df_t.columns:
        if any(x in c.lower() for x in ["created", "tanggal", "date", "tgl", "waktu"]):
            date_col = c
            break
    if date_col is None:
        return pd.DataFrame()

    df_t["_tgl"] = pd.to_datetime(df_t[date_col], errors="coerce").dt.normalize()
    daily_tiket = (
        df_t[df_t["_tgl"].notna()]
        .groupby("_tgl")["Nominal"]
        .sum()
        .reset_index()
        .rename(columns={"_tgl": "Tanggal", "Nominal": "Total Tiket (Rp)"})
    )

    if df_bank is None or len(df_bank) == 0:
        # Tidak ada data bank — tampilkan semua hari sebagai tidak terverifikasi
        daily_tiket["Total Bank Masuk (Rp)"] = None
        daily_tiket["Selisih (Rp)"]          = None
        daily_tiket["Status"]                = "⚠️ Belum ada data bank"
        return daily_tiket

    # Hitung uang masuk Prepaid / OnUs di rekening koran
    # Kategori Prepaid mencakup: "Prepaid", "Cash" (BCA tanpa mrc), "NON BCA" (BRI Pinbuk), "NON MANDIRI"
    PREPAID_CATEGORIES = {"Prepaid", "Cash"}
    df_b = df_bank[df_bank["Kategori"].isin(PREPAID_CATEGORIES)].copy()
    df_b["_tgl"] = pd.to_datetime(df_b["Tanggal"], errors="coerce").dt.normalize()
    daily_bank = (
        df_b[df_b["_tgl"].notna()]
        .groupby("_tgl")["Nominal"]
        .sum()
        .reset_index()
        .rename(columns={"_tgl": "Tanggal", "Nominal": "Total Bank Masuk (Rp)"})
    )

    merged = pd.merge(daily_tiket, daily_bank, on="Tanggal", how="left")
    merged["Total Bank Masuk (Rp)"] = merged["Total Bank Masuk (Rp)"].fillna(0)
    merged["Selisih (Rp)"] = merged["Total Bank Masuk (Rp)"] - merged["Total Tiket (Rp)"]

    def status_fn(row):
        if row["Total Bank Masuk (Rp)"] == 0:
            return "🔴 Tidak ada masuk di bank"
        elif row["Selisih (Rp)"] < 0:
            return "🟠 Kurang Settlement"
        elif row["Selisih (Rp)"] > 0:
            return "🔵 Lebih Settlement"
        else:
            return "✅ Sesuai"

    merged["Status"] = merged.apply(status_fn, axis=1)
    merged["Tanggal"] = merged["Tanggal"].dt.strftime("%d %b %Y")
    return merged


# ── Tabs ──────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "   📂 Upload File   ",
    "   🔍 Rekonsiliasi Order ID   ",
    "   📊 Ringkasan   ",
    "   🏦 Data Bank   ",
])

# ══════════════════════════════════════════════════════════════
# TAB 1 — Upload File
# ══════════════════════════════════════════════════════════════
with tab1:
    st.markdown("### Upload File Excel")
    st.markdown("Pilih file Excel untuk **Tiket Detail**, **Settlement Espay**, dan opsional **Rekening Koran Bank**.")

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
    st.markdown("#### 🏦 Rekening Koran Bank _(opsional)_")
    st.markdown("Upload rekening koran untuk memverifikasi uang masuk dari settlement ESPAY maupun Prepaid.")

    # Format note rekening koran
    st.markdown("""
<div class="warn-box">
<b>📋 Format File Rekening Koran yang Didukung:</b><br>
File Excel dengan kolom berikut (nama kolom fleksibel, urutan sesuai):<br><br>
<table style="width:100%; border-collapse:collapse; font-size:13px;">
  <tr style="background:#FED7AA">
    <th style="padding:6px 10px; text-align:left; border:1px solid #FDBA74">Kolom</th>
    <th style="padding:6px 10px; text-align:left; border:1px solid #FDBA74">Nama yang Dikenali</th>
    <th style="padding:6px 10px; text-align:left; border:1px solid #FDBA74">Keterangan</th>
  </tr>
  <tr>
    <td style="padding:6px 10px; border:1px solid #FDBA74"><b>Tanggal</b></td>
    <td style="padding:6px 10px; border:1px solid #FDBA74">Tanggal / Date / Tgl</td>
    <td style="padding:6px 10px; border:1px solid #FDBA74">Tanggal transaksi</td>
  </tr>
  <tr style="background:#FFF7ED">
    <td style="padding:6px 10px; border:1px solid #FDBA74"><b>Keterangan</b></td>
    <td style="padding:6px 10px; border:1px solid #FDBA74">Keterangan / Deskripsi / Uraian / Remark</td>
    <td style="padding:6px 10px; border:1px solid #FDBA74">Deskripsi transaksi (digunakan untuk klasifikasi)</td>
  </tr>
  <tr>
    <td style="padding:6px 10px; border:1px solid #FDBA74"><b>Nominal</b></td>
    <td style="padding:6px 10px; border:1px solid #FDBA74">Kredit / Masuk / Nominal / Jumlah / Amount</td>
    <td style="padding:6px 10px; border:1px solid #FDBA74">Nilai uang <b>masuk</b> (kredit). Baris dengan nominal 0 diabaikan.</td>
  </tr>
</table>
<br>
<b>Logika Klasifikasi Transaksi:</b><br>
• Keterangan mengandung <b>"mrc"</b> → dicek ke data Settlement Espay → <b>ESPAY</b> (jika ditemukan)<br>
• BRI: Keterangan mengandung <b>"Pinbuk"</b> → Transfer NON BCA &nbsp;|&nbsp; Tidak ada keduanya → <b>Prepaid (OnUs)</b><br>
• BCA: Keterangan mengandung <b>"KR OTOMATIS"</b> + <b>"MID"</b> → Transfer BCA &nbsp;|&nbsp; Tidak ada → Cash<br>
• Mandiri: Keterangan mengandung <b>"SWITCHING / EDC / MTRANSFER"</b> → Transfer NON MANDIRI &nbsp;|&nbsp; Tidak ada → <b>Prepaid (OnUs)</b>
</div>
""", unsafe_allow_html=True)

    st.markdown("")
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        st.markdown("**🏦 Rekening Koran BRI**")
        st.caption("Sheet pertama atau sheet yang mengandung kata 'koran'/'mutasi'")
        file_bri = st.file_uploader("Upload Rekening Koran BRI", type=["xlsx", "xls"], key="bri")
    with bc2:
        st.markdown("**🏦 Rekening Koran BCA**")
        st.caption("Sheet pertama atau sheet yang mengandung kata 'koran'/'mutasi'")
        file_bca = st.file_uploader("Upload Rekening Koran BCA", type=["xlsx", "xls"], key="bca")
    with bc3:
        st.markdown("**🏦 Rekening Koran Mandiri**")
        st.caption("Sheet pertama atau sheet yang mengandung kata 'koran'/'mutasi'")
        file_mandiri = st.file_uploader("Upload Rekening Koran Mandiri", type=["xlsx", "xls"], key="mandiri")

    st.markdown("---")

    st.markdown("""
<div class="info-box">
<b>Cara Pakai:</b><br>
1. Upload file Excel Rekonsiliasi Cabang (untuk Tiket Detail).<br>
2. Upload file yang sama atau file lain untuk Settlement Espay.<br>
3. (Opsional) Upload Rekening Koran BRI, BCA, dan/atau Mandiri untuk verifikasi uang masuk.<br>
4. Klik <b>Proses Rekonsiliasi</b> — hasil muncul di tab Rekonsiliasi Order ID, Ringkasan, dan Data Bank.
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
                    # ── Tiket Detail ─────────────────────────────────────
                    df_t = pd.read_excel(file_tiket, sheet_name="Tiket Detail", dtype={"Order ID": str})
                    nom_col = df_t.columns[24]
                    df_t = df_t.rename(columns={nom_col: "Nominal"})
                    df_t = df_t[df_t["St Bayar"].astype(str).str.lower().str.strip() == "paid"].copy()
                    df_t["Order ID"] = df_t["Order ID"].astype(str).str.strip()
                    df_t["Bank"] = df_t["Bank"].astype(str).str.strip()

                    # Pisahkan ESPAY vs Prepaid
                    # ESPAY: Bank = "ESPAY" atau "ESPAY MANDIRI" / "MANDIRI ESPAY"
                    espay_mask = (
                        df_t["Bank"].str.upper().str.contains("ESPAY", na=False)
                    )
                    df_espay       = df_t[espay_mask].copy()
                    df_prepaid_raw = df_t[~espay_mask].copy()

                    # Ringkasan prepaid per Bank
                    prepaid_grp = (df_prepaid_raw.groupby("Bank")
                                   .agg(jumlah_tiket=("Nominal", "count"),
                                        nominal=("Nominal", "sum"))
                                   .reset_index()
                                   .rename(columns={"jumlah_tiket": "Jml Tiket", "nominal": "Nominal (Rp)"}))
                    st.session_state.df_prepaid     = prepaid_grp
                    st.session_state.df_prepaid_raw = df_prepaid_raw

                    # ── Settlement Espay ──────────────────────────────────
                    df_s = pd.read_excel(file_settle, sheet_name="Settlement Espay", dtype={"Order Id": str})
                    tipe_col = df_s.columns[24]
                    df_s = df_s.rename(columns={tipe_col: "Tipe", "Order Id": "Order ID"})
                    df_s["Order ID"] = df_s["Order ID"].astype(str).str.strip()

                    # Filter periode: rentang tanggal dari tiket ESPAY
                    df_espay["Created"] = pd.to_datetime(df_espay["Created"], errors="coerce")
                    df_s["Transaction Date"] = pd.to_datetime(df_s["Transaction Date"], errors="coerce")

                    periode_min = df_espay["Created"].min().normalize()
                    periode_max = (df_espay["Created"].max().normalize()
                                   + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))

                    df_s_filtered = df_s[
                        (df_s["Transaction Date"] >= periode_min) &
                        (df_s["Transaction Date"] <= periode_max)
                    ].copy()

                    st.session_state.periode_min     = periode_min
                    st.session_state.periode_max     = periode_max
                    st.session_state.settle_total    = len(df_s)
                    st.session_state.settle_filtered = len(df_s_filtered)
                    st.session_state.df_settle_raw   = df_s_filtered.copy()

                    # ── Grouping rekonsiliasi ─────────────────────────────
                    tiket_grp = (df_espay.groupby("Order ID")
                                 .agg(jumlah_tiket=("Nominal", "count"),
                                      nominal_tiket=("Nominal", "sum"))
                                 .reset_index())

                    settle_grp = (df_s_filtered.groupby("Order ID")
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

                    st.session_state.df_rekon  = df_m
                    st.session_state.df_espay  = df_espay
                    st.session_state.processed = True

                    # ── Proses Rekening Koran Bank ────────────────────────
                    # Simpan df_s agar bisa digunakan saat klasifikasi bank
                    espay_for_bank = df_s.copy()

                    df_bank_bri     = None
                    df_bank_bca     = None
                    df_bank_mandiri = None

                    if file_bri:
                        df_bank_bri = process_bank_file(file_bri, "BRI", espay_for_bank)
                        st.session_state.df_bank_bri = df_bank_bri

                    if file_bca:
                        df_bank_bca = process_bank_file(file_bca, "BCA", espay_for_bank)
                        st.session_state.df_bank_bca = df_bank_bca

                    if file_mandiri:
                        df_bank_mandiri = process_bank_file(file_mandiri, "MANDIRI", espay_for_bank)
                        st.session_state.df_bank_mandiri = df_bank_mandiri

                    if file_bri or file_bca or file_mandiri:
                        st.session_state.bank_processed = True

                    # Summary counts
                    match_c   = (df_m["Status"] == "Match").sum()
                    no_set_c  = (df_m["Status"] == "Tidak di Settlement").sum()
                    no_tkt_c  = (df_m["Status"] == "Tidak di Tiket Detail").sum()
                    prepaid_c = df_prepaid_raw["Order ID"].nunique()

                    bank_info = ""
                    if file_bri and df_bank_bri is not None:
                        bank_info += f"  |  BRI: {len(df_bank_bri):,} baris"
                    if file_bca and df_bank_bca is not None:
                        bank_info += f"  |  BCA: {len(df_bank_bca):,} baris"
                    if file_mandiri and df_bank_mandiri is not None:
                        bank_info += f"  |  Mandiri: {len(df_bank_mandiri):,} baris"

                    st.success(
                        f"✅ Selesai — {len(df_m):,} unique Order ID ESPAY  |  "
                        f"Match: {match_c:,}  |  "
                        f"Tidak di Settlement: {no_set_c:,}  |  "
                        f"Tidak di Tiket Detail: {no_tkt_c:,}  |  "
                        f"Prepaid (non-ESPAY): {prepaid_c:,} Order ID"
                        + bank_info
                    )
                    st.info(
                        f"📅 Periode tiket: **{periode_min.strftime('%d %b %Y')}** s/d "
                        f"**{df_espay['Created'].max().strftime('%d %b %Y')}**  |  "
                        f"Settlement dibaca: **{len(df_s_filtered):,}** dari {len(df_s):,} baris total"
                    )
                    st.info("Buka tab **Rekonsiliasi Order ID**, **Ringkasan**, atau **Data Bank** untuk melihat hasil.")

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
        df       = st.session_state.df_rekon
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

        # Info periode
        if st.session_state.get("periode_min") is not None:
            p_min = st.session_state.periode_min
            p_max = st.session_state.periode_max
            s_tot = st.session_state.settle_total
            s_fil = st.session_state.settle_filtered
            st.markdown(
                f'<div class="info-box">📅 <b>Periode Tiket Detail (ESPAY):</b> '
                f'{p_min.strftime("%d %b %Y")} s/d {p_max.strftime("%d %b %Y")} &nbsp;|&nbsp; '
                f'Settlement digunakan: <b>{s_fil:,}</b> dari {s_tot:,} baris total</div>',
                unsafe_allow_html=True
            )
            st.markdown("")

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

        # ── Tabel Kurang Settlement Prepaid ───────────────────
        df_prepaid_raw = st.session_state.df_prepaid_raw
        df_bank_bri    = st.session_state.df_bank_bri
        df_bank_bca    = st.session_state.df_bank_bca
        df_bank_mandiri = st.session_state.df_bank_mandiri

        st.markdown("#### 🟠 Kurang Settlement Prepaid (Perbandingan Tiket vs Uang Masuk Bank)")

        if df_prepaid_raw is not None and len(df_prepaid_raw):
            # Gabungkan semua data bank
            bank_frames = []
            if df_bank_bri is not None and len(df_bank_bri):
                bank_frames.append(df_bank_bri)
            if df_bank_bca is not None and len(df_bank_bca):
                bank_frames.append(df_bank_bca)
            if df_bank_mandiri is not None and len(df_bank_mandiri):
                bank_frames.append(df_bank_mandiri)
            df_bank_all = pd.concat(bank_frames, ignore_index=True) if bank_frames else None

            shortfall_df = detect_prepaid_shortfall(df_prepaid_raw, df_bank_all)

            if len(shortfall_df):
                # Warna status
                def color_shortfall(val):
                    if "Kurang" in str(val):
                        return "background-color: #FFF7ED; color: #C2410C"
                    elif "Tidak ada" in str(val):
                        return "background-color: #FEF2F2; color: #DC2626"
                    elif "Lebih" in str(val):
                        return "background-color: #EFF6FF; color: #1D4ED8"
                    elif "Sesuai" in str(val):
                        return "background-color: #F0FDF4; color: #15803D"
                    return ""

                fmt_dict = {}
                if "Total Tiket (Rp)" in shortfall_df.columns:
                    fmt_dict["Total Tiket (Rp)"] = "{:,.0f}"
                if "Total Bank Masuk (Rp)" in shortfall_df.columns:
                    fmt_dict["Total Bank Masuk (Rp)"] = lambda x: f"{x:,.0f}" if pd.notna(x) else "—"
                if "Selisih (Rp)" in shortfall_df.columns:
                    fmt_dict["Selisih (Rp)"] = lambda x: f"{x:,.0f}" if pd.notna(x) else "—"

                styled_sf = shortfall_df.style.map(color_shortfall, subset=["Status"]).format(fmt_dict)
                st.dataframe(styled_sf, use_container_width=True, hide_index=True,
                             height=min(400, 35 * len(shortfall_df) + 38))

                # Ringkasan shortfall
                if "Selisih (Rp)" in shortfall_df.columns:
                    total_kurang = shortfall_df[shortfall_df["Status"].str.contains("Kurang", na=False)]["Selisih (Rp)"].sum()
                    hari_kurang  = (shortfall_df["Status"].str.contains("Kurang", na=False)).sum()
                    hari_kosong  = (shortfall_df["Status"].str.contains("Tidak ada", na=False)).sum()
                    if total_kurang != 0 or hari_kosong:
                        st.markdown(f"""
<div class="warn-box">
⚠️ <b>{hari_kurang} hari</b> dengan kurang settlement prepaid &nbsp;|&nbsp;
<b>{hari_kosong} hari</b> tanpa uang masuk di bank &nbsp;|&nbsp;
Total kurang: <b>Rp {abs(total_kurang):,.0f}</b><br>
<small>Catatan: Nilai settlement prepaid mengikuti data cut-off di Tiket Detail.
Pastikan rekening koran yang diupload mencakup seluruh periode.</small>
</div>""", unsafe_allow_html=True)
                    else:
                        st.success("✅ Semua hari prepaid telah ter-settlement dengan baik.")
            else:
                st.warning("⚠️ Tidak dapat mendeteksi kolom tanggal di data Prepaid. Pastikan kolom 'Created' tersedia di Tiket Detail.")
        else:
            st.info("Tidak ada transaksi Prepaid di Tiket Detail.")

        # ── Tabel 1: Tiket ESPAY tidak di Settlement ──────────
        st.markdown("---")
        hilang_set_df = df[df["Status"] == "Tidak di Settlement"][
            ["Order ID", "Nominal Tiket (Rp)", "Jml Tiket"]].copy()

        df_espay_ss = st.session_state.df_espay
        if df_espay_ss is not None:
            golongan_map = (df_espay_ss.groupby("Order ID")["Golongan"]
                            .apply(lambda x: ", ".join(x.astype(str).unique()))
                            .reset_index()
                            .rename(columns={"Golongan": "Golongan"}))
            hilang_set_df = hilang_set_df.merge(golongan_map, on="Order ID", how="left")
            hilang_set_df = hilang_set_df[["Order ID", "Golongan", "Jml Tiket", "Nominal Tiket (Rp)"]]

        st.markdown(f"#### 🔴 Tiket ESPAY Tidak Ada di Settlement ({no_set_c:,} Order ID)")
        if len(hilang_set_df):
            st.dataframe(
                hilang_set_df.style.format({"Nominal Tiket (Rp)": "{:,.0f}"}),
                use_container_width=True,
                hide_index=True,
                height=min(400, 35 * len(hilang_set_df) + 38),
            )
        else:
            st.success("✅ Tidak ada tiket ESPAY yang hilang di Settlement.")

        # ── Tabel 2: Settlement tidak di Tiket Detail ─────────
        st.markdown("---")
        lebih_set_df = df[df["Status"] == "Tidak di Tiket Detail"][["Order ID"]].copy()

        if df_s_raw is not None and len(lebih_set_df):
            settle_detail = (df_s_raw[df_s_raw["Order ID"].isin(lebih_set_df["Order ID"])]
                             [["Order ID", "Transaction Date", "BANK", "Tipe", "Amount"]]
                             .copy())
            settle_detail["Transaction Date"] = pd.to_datetime(
                settle_detail["Transaction Date"], errors="coerce"
            ).dt.strftime("%d %b %Y %H:%M")
            lebih_detail_df = (settle_detail.groupby("Order ID")
                               .agg(
                                   Transaction_Date=("Transaction Date", "first"),
                                   Bank=("BANK", lambda x: x.mode().iloc[0] if len(x) else ""),
                                   Tipe=("Tipe", lambda x: x.mode().iloc[0] if len(x) else ""),
                                   Amount=("Amount", "sum"),
                               )
                               .reset_index()
                               .rename(columns={
                                   "Transaction_Date": "Transaction Date",
                                   "Amount": "Amount (Rp)",
                               }))
        else:
            lebih_detail_df = pd.DataFrame(columns=["Order ID", "Transaction Date", "Bank", "Tipe", "Amount (Rp)"])

        st.markdown(f"#### 🟡 Settlement Tidak Ada di Tiket Detail ({no_tkt_c:,} Order ID)")
        if len(lebih_detail_df):
            st.dataframe(
                lebih_detail_df.style.format({"Amount (Rp)": "{:,.0f}"}),
                use_container_width=True,
                hide_index=True,
                height=min(400, 35 * len(lebih_detail_df) + 38),
            )
        else:
            st.success("✅ Tidak ada Order ID di Settlement yang tidak ditemukan di Tiket Detail.")

# ══════════════════════════════════════════════════════════════
# TAB 4 — Data Bank
# ══════════════════════════════════════════════════════════════
with tab4:
    df_bank_bri     = st.session_state.df_bank_bri
    df_bank_bca     = st.session_state.df_bank_bca
    df_bank_mandiri = st.session_state.df_bank_mandiri

    has_bank = (
        (df_bank_bri     is not None and len(df_bank_bri)     > 0) or
        (df_bank_bca     is not None and len(df_bank_bca)     > 0) or
        (df_bank_mandiri is not None and len(df_bank_mandiri) > 0)
    )

    if not has_bank:
        st.info("Belum ada data rekening koran. Upload file Rekening Koran BRI/BCA di tab **Upload File** lalu klik **Proses Rekonsiliasi**.")
    else:
        st.markdown("### 🏦 Analisis Uang Masuk Rekening Koran Bank")

        # Gabungkan semua data bank
        bank_frames = []
        if df_bank_bri is not None and len(df_bank_bri):
            bank_frames.append(df_bank_bri)
        if df_bank_bca is not None and len(df_bank_bca):
            bank_frames.append(df_bank_bca)
        if df_bank_mandiri is not None and len(df_bank_mandiri):
            bank_frames.append(df_bank_mandiri)
        df_all = pd.concat(bank_frames, ignore_index=True)

        # ── Ringkasan Total per Bank ──────────────────────────
        st.markdown("#### 📊 Ringkasan Total Uang Masuk per Bank")

        summary_bank = (df_all.groupby(["Bank", "Kategori"])["Nominal"]
                        .agg(["sum", "count"])
                        .reset_index()
                        .rename(columns={"sum": "Total (Rp)", "count": "Jumlah Transaksi"}))

        for bank_name in df_all["Bank"].unique():
            df_b = df_all[df_all["Bank"] == bank_name]
            total_masuk = df_b["Nominal"].sum()

            # Kategori ESPAY: baris yang berhasil di-match ke Settlement (mrc found)
            # → kategori berisi nilai BANK dari col AA: "NON BCA", "BCA", dll (bukan "ESPAY")
            # Deteksi ESPAY = baris yang punya Order Ref (mrc match)
            espay_mask    = df_b["Order Ref"].notna()
            prepaid_mask  = df_b["Kategori"].isin(["Prepaid", "Cash"])
            transfer_mask = df_b["Kategori"].isin(["NON BCA", "NON MANDIRI", "BCA"])

            total_espay    = df_b[espay_mask]["Nominal"].sum()
            total_prepaid  = df_b[prepaid_mask]["Nominal"].sum()
            total_transfer = df_b[transfer_mask & ~espay_mask]["Nominal"].sum()
            total_lain     = df_b[~espay_mask & ~prepaid_mask & ~transfer_mask]["Nominal"].sum()
            n_espay   = espay_mask.sum()
            n_prepaid = prepaid_mask.sum()

            st.markdown(f"##### 🏦 {bank_name}")
            bk1, bk2, bk3, bk4 = st.columns(4)
            with bk1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Total Uang Masuk</div>
                    <div class="metric-value" style="color:#2563EB; font-size:15px">Rp {total_masuk:,.0f}</div>
                </div>""", unsafe_allow_html=True)
            with bk2:
                st.markdown(f"""
                <div class="green-card">
                    <div style="font-size:13px;color:#166534">ESPAY (mrc)</div>
                    <div style="font-size:16px;font-weight:700;color:#16A34A">Rp {total_espay:,.0f}</div>
                    <div style="font-size:12px;color:#6B7280">{int(n_espay):,} transaksi</div>
                </div>""", unsafe_allow_html=True)
            with bk3:
                st.markdown(f"""
                <div class="purple-card">
                    <div style="font-size:13px;color:#5B21B6">Prepaid (OnUs)</div>
                    <div style="font-size:16px;font-weight:700;color:#7C3AED">Rp {total_prepaid:,.0f}</div>
                    <div style="font-size:12px;color:#6B7280">{int(n_prepaid):,} transaksi</div>
                </div>""", unsafe_allow_html=True)
            with bk4:
                st.markdown(f"""
                <div class="bank-card">
                    <div style="font-size:13px;color:#1E40AF">Transfer / Lainnya</div>
                    <div style="font-size:16px;font-weight:700;color:#2563EB">Rp {total_transfer + total_lain:,.0f}</div>
                    <div style="font-size:12px;color:#6B7280">&nbsp;</div>
                </div>""", unsafe_allow_html=True)

            st.markdown("")

        # ── Tabel ringkasan per bank & kategori ──────────────
        st.markdown("#### 📋 Tabel Ringkasan per Bank & Kategori")
        summary_pivot = summary_bank.copy()
        summary_pivot["Total (Rp)"] = summary_pivot["Total (Rp)"].apply(lambda x: f"Rp {x:,.0f}")
        st.dataframe(summary_pivot, use_container_width=True, hide_index=True)

        # ── Grafik harian per bank ────────────────────────────
        st.markdown("---")
        st.markdown("#### 📈 Grafik Uang Masuk Harian — ESPAY vs Prepaid")

        import json

        for bank_name in df_all["Bank"].unique():
            df_b = df_all[df_all["Bank"] == bank_name].copy()
            df_b["Tgl"] = pd.to_datetime(df_b["Tanggal"], errors="coerce").dt.normalize()
            df_b = df_b[df_b["Tgl"].notna()]

            # Harian per kategori
            espay_daily = (df_b[df_b["Kategori"] == "ESPAY"]
                           .groupby("Tgl")["Nominal"].sum()
                           .reset_index()
                           .rename(columns={"Nominal": "ESPAY"}))
            prepaid_daily = (df_b[df_b["Kategori"].isin(["Prepaid", "Cash"])]
                             .groupby("Tgl")["Nominal"].sum()
                             .reset_index()
                             .rename(columns={"Nominal": "Prepaid/OnUs"}))

            all_dates = pd.date_range(df_b["Tgl"].min(), df_b["Tgl"].max())
            daily = pd.DataFrame({"Tgl": all_dates})
            daily = daily.merge(espay_daily, on="Tgl", how="left")
            daily = daily.merge(prepaid_daily, on="Tgl", how="left")
            daily = daily.fillna(0)
            daily["Label"] = daily["Tgl"].dt.strftime("%d %b")

            labels      = daily["Label"].tolist()
            espay_vals  = daily["ESPAY"].tolist()
            prepaid_vals = daily["Prepaid/OnUs"].tolist()

            chart_html = f"""
<div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:16px;margin-bottom:24px">
  <h4 style="margin:0 0 12px;color:#1E293B;font-size:15px">🏦 {bank_name} — Uang Masuk Harian</h4>
  <canvas id="chart_{bank_name}" height="200"></canvas>
</div>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script>
(function() {{
  var ctx = document.getElementById('chart_{bank_name}').getContext('2d');
  new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: {json.dumps(labels)},
      datasets: [
        {{
          label: 'ESPAY (mrc)',
          data: {json.dumps(espay_vals)},
          backgroundColor: 'rgba(22,163,74,0.7)',
          borderColor: '#16A34A',
          borderWidth: 1,
          borderRadius: 4,
        }},
        {{
          label: 'Prepaid (OnUs)',
          data: {json.dumps(prepaid_vals)},
          backgroundColor: 'rgba(124,58,237,0.7)',
          borderColor: '#7C3AED',
          borderWidth: 1,
          borderRadius: 4,
        }}
      ]
    }},
    options: {{
      responsive: true,
      plugins: {{
        legend: {{ position: 'top' }},
        tooltip: {{
          callbacks: {{
            label: function(ctx) {{
              return ctx.dataset.label + ': Rp ' + ctx.parsed.y.toLocaleString('id-ID');
            }}
          }}
        }}
      }},
      scales: {{
        x: {{ stacked: false }},
        y: {{
          stacked: false,
          ticks: {{
            callback: function(v) {{
              if(v >= 1000000) return 'Rp ' + (v/1000000).toFixed(1) + ' Jt';
              return 'Rp ' + v.toLocaleString('id-ID');
            }}
          }}
        }}
      }}
    }}
  }});
}})();
</script>
"""
            st.components.v1.html(chart_html, height=320)

        # ── Detail transaksi per bank ─────────────────────────
        st.markdown("---")
        st.markdown("#### 🔎 Detail Transaksi Rekening Koran")

        sel_bank = st.selectbox(
            "Pilih Bank",
            options=df_all["Bank"].unique().tolist(),
            key="bank_detail_select"
        )
        sel_kat = st.selectbox(
            "Filter Kategori",
            options=["Semua"] + sorted(df_all["Kategori"].unique().tolist()),
            key="kat_detail_select"
        )

        df_detail = df_all[df_all["Bank"] == sel_bank].copy()
        if sel_kat != "Semua":
            df_detail = df_detail[df_detail["Kategori"] == sel_kat]

        df_detail["Tanggal"] = pd.to_datetime(df_detail["Tanggal"], errors="coerce").dt.strftime("%d %b %Y")

        disp_cols = ["Tanggal", "Keterangan", "Nominal", "Kategori", "Tipe", "Order Ref"]
        df_disp = df_detail[[c for c in disp_cols if c in df_detail.columns]].copy()

        def color_kat(val):
            if val == "ESPAY":
                return "background-color:#F0FDF4; color:#16A34A"
            elif val in ("Prepaid", "Cash"):
                return "background-color:#F5F3FF; color:#7C3AED"
            elif val == "Transfer":
                return "background-color:#EFF6FF; color:#1D4ED8"
            return ""

        styled_det = df_disp.style.map(color_kat, subset=["Kategori"]).format({"Nominal": "{:,.0f}"})
        st.dataframe(styled_det, use_container_width=True, hide_index=True,
                     height=min(500, 35 * len(df_disp) + 38))

        # Download detail
        buf_bank = io.BytesIO()
        with pd.ExcelWriter(buf_bank, engine="openpyxl") as w:
            for bank_name in df_all["Bank"].unique():
                df_all[df_all["Bank"] == bank_name].drop(columns=[], errors="ignore").to_excel(
                    w, index=False, sheet_name=f"Bank {bank_name}"[:31]
                )
        st.download_button(
            "⬇️ Download Data Bank (Excel)",
            data=buf_bank.getvalue(),
            file_name=f"data_bank_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
