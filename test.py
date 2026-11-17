"""SiCupang Akuntansi - aplikasi akuntansi sederhana (cash basis).

Seluruh program digabung dalam satu file. Jalankan dengan:  streamlit run app.py
"""

from datetime import date, timedelta
import io
import uuid

from openpyxl.utils import get_column_letter
from supabase import create_client
import pandas as pd
import streamlit as st

# set_page_config harus menjadi perintah Streamlit pertama yang dijalankan
st.set_page_config(page_title="SiCupang Akuntansi", page_icon="🐟", layout="wide",
                   initial_sidebar_state="expanded")


# ####################################################################
# BAGIAN 1. KONSTANTA & CHART OF ACCOUNTS
# (asal: COA.py)
# ####################################################################

# =====================================================================
# 1. DAFTAR AKUN (Chart of Accounts)
# =====================================================================
# Format:  "kode": {"nama": ..., "tipe": ..., "normal": "D" atau "K"}
#   tipe   : Aset / Liabilitas / Ekuitas / Pendapatan / Beban
#   normal : sisi saldo normal akun. D = Debit, K = Kredit
# Akun di bawah ini adalah akun BAWAAN. Akun tambahan dari menu "Tambah COA"
# disimpan di database (tabel coa_custom) dan digabung lewat semua_akun().
AKUN_DASAR = {
    "1-100": {"nama": "Kas",                   "tipe": "Aset",       "normal": "D"},
    "1-101": {"nama": "Kas di Bank",           "tipe": "Aset",       "normal": "D"},
    "1-102": {"nama": "Piutang",               "tipe": "Aset",       "normal": "D"},
    "1-103": {"nama": "Persediaan Ikan",       "tipe": "Aset",       "normal": "D"},
    "1-104": {"nama": "Persediaan Pakan",      "tipe": "Aset",       "normal": "D"},
    "1-105": {"nama": "Persediaan Obat",       "tipe": "Aset",       "normal": "D"},
    "1-106": {"nama": "Persediaan Lainnya",    "tipe": "Aset",       "normal": "D"},
    "1-107": {"nama": "Perlengkapan",          "tipe": "Aset",       "normal": "D"},
    "1-108": {"nama": "Peralatan",             "tipe": "Aset",       "normal": "D"},
    "1-109": {"nama": "Akumulasi Penyusutan Peralatan", "tipe": "Aset", "normal": "K"},
    "2-101": {"nama": "Utang Bank",            "tipe": "Liabilitas", "normal": "K"},
    "2-102": {"nama": "Utang Orang",           "tipe": "Liabilitas", "normal": "K"},
    "3-101": {"nama": "Modal",                 "tipe": "Ekuitas",    "normal": "K"},
    "3-102": {"nama": "Iktisar Laba Rugi",     "tipe": "Ekuitas",    "normal": "K"},
    "3-103": {"nama": "Prive",                 "tipe": "Ekuitas",    "normal": "D"},
    "4-101": {"nama": "Penjualan",             "tipe": "Pendapatan", "normal": "K"},
    "4-102": {"nama": "Penjualan Lainnya",     "tipe": "Pendapatan", "normal": "K"},   # lama, tidak dipakai transaksi baru
    "5-101": {"nama": "Beban Pembelian Ikan",  "tipe": "Beban",      "normal": "D"},   # lama
    "5-102": {"nama": "Beban Pakan",           "tipe": "Beban",      "normal": "D"},   # lama
    "5-103": {"nama": "Beban Obat",            "tipe": "Beban",      "normal": "D"},   # lama
    "5-104": {"nama": "Beban Perlengkapan",    "tipe": "Beban",      "normal": "D"},
    "5-105": {"nama": "Beban Lain-lain",       "tipe": "Beban",      "normal": "D"},
    "5-106": {"nama": "Pembelian",             "tipe": "Beban",      "normal": "D"},
    "5-107": {"nama": "Beban Penyusutan Peralatan", "tipe": "Beban", "normal": "D"},
}

# Kode akun yang dipakai di kode program
KAS_TANGAN = "1-100"
KAS_BANK = "1-101"
KAS = KAS_BANK                      # nama lama, dipertahankan
PIUTANG = "1-102"
PERLENGKAPAN = "1-107"
PERALATAN = "1-108"
AKUM_PENYUSUTAN = "1-109"
MODAL = "3-101"
IKHTISAR = "3-102"
PRIVE = "3-103"
PENJUALAN = "4-101"
BEBAN_PERLENGKAPAN = "5-104"
PEMBELIAN = "5-106"
BEBAN_PENYUSUTAN = "5-107"

# Akun persediaan per kategori (untuk penyesuaian / stock opname)
AKUN_PERSEDIAAN = {"Ikan": "1-103", "Pakan": "1-104", "Obat": "1-105", "Lainnya": "1-106"}

# Akun yang ditampilkan sebagai Harga Pokok Penjualan di laporan laba rugi
HPP_AKUN = ("5-101", "5-102", "5-103", PEMBELIAN)

# Kategori barang yang dijual / dibeli sebagai persediaan
KATEGORI_BARANG = ["Ikan", "Pakan", "Obat"]

# Pilihan kas di form: "tulisan di layar": "kode akun"
KAS_PILIHAN = {"Kas": KAS_TANGAN, "Kas di Bank": KAS_BANK}

TIPE_AKUN = ["Aset", "Liabilitas", "Ekuitas", "Pendapatan", "Beban"]
NORMAL_BAWAAN = {"Aset": "D", "Beban": "D", "Liabilitas": "K", "Ekuitas": "K", "Pendapatan": "K"}
AWAL_KODE = {"Aset": "1", "Liabilitas": "2", "Ekuitas": "3", "Pendapatan": "4", "Beban": "5"}


def semua_akun():
    """Akun bawaan + akun tambahan milik user yang sedang login, urut menurut kode."""
    gabung = {**AKUN_DASAR, **st.session_state.get("akun_custom", {})}
    return dict(sorted(gabung.items()))


def pilihan_pembelian():
    """Dropdown halaman Pembelian: {"tulisan": "kode akun yang di-DEBIT"}.
    Ikan/Pakan/Obat -> akun Pembelian. Semua akun Beban (termasuk tambahan dari menu
    Tambah COA) otomatis ikut."""
    pilihan = {k: PEMBELIAN for k in KATEGORI_BARANG}
    pilihan["Perlengkapan"] = BEBAN_PERLENGKAPAN
    pilihan["Peralatan (dicatat sebagai aset)"] = PERALATAN
    dikecualikan = {PEMBELIAN, BEBAN_PERLENGKAPAN, BEBAN_PENYUSUTAN, "5-101", "5-102", "5-103"}
    for kode, info in semua_akun().items():
        if info["tipe"] == "Beban" and kode not in dikecualikan:
            pilihan[info["nama"]] = kode
    return pilihan


# ####################################################################
# BAGIAN 2. PERHITUNGAN & LAPORAN AKUNTANSI
# (asal: akuntansi.py)
# ####################################################################

# =====================================================================
# 4. PERHITUNGAN AKUNTANSI
# =====================================================================
def saldo_per_akun(df):
    """Hitung saldo tiap akun. Saldo dihitung sesuai sisi normalnya:
       akun normal Debit  -> saldo = debit - kredit
       akun normal Kredit -> saldo = kredit - debit
    Hasil: {"1-101": 5000000, ...}"""
    g = df.groupby("kode_akun")[["debit", "kredit"]].sum()
    hasil = {}
    for kode, info in semua_akun().items():
        if kode in g.index:
            d, k = int(g.at[kode, "debit"]), int(g.at[kode, "kredit"])
        else:
            d = k = 0
        hasil[kode] = d - k if info["normal"] == "D" else k - d
    return hasil


def total_tipe(saldo, tipe):
    """Jumlahkan saldo semua akun bertipe tertentu (mis. 'Pendapatan')."""
    akun = semua_akun()
    return sum(s for kode, s in saldo.items() if akun[kode]["tipe"] == tipe)


def filter_periode(df, awal=None, akhir=None):
    """Ambil baris jurnal pada rentang tanggal tertentu."""
    if awal is not None:
        df = df[df["tanggal"] >= awal]
    if akhir is not None:
        df = df[df["tanggal"] <= akhir]
    return df


def tanpa_penutup(df):
    """Buang jurnal penutup. Dipakai untuk laporan laba rugi / dashboard, supaya
    pendapatan & beban tetap terlihat walau periodenya sudah ditutup."""
    return df[df["jenis"] != "penutup"]


def nilai_persediaan_iktisar(df):
    """Persediaan awal & akhir yang dipindahkan lewat Iktisar Laba Rugi saat penyesuaian.
    Debit iktisar = persediaan awal, Kredit iktisar = persediaan akhir."""
    ik = df[(df["kode_akun"] == IKHTISAR) & (df["jenis"] == "penyesuaian")]
    return int(ik["debit"].sum()), int(ik["kredit"].sum())


def hitung_laba(df):
    """(pendapatan, beban, laba). df = jurnal tanpa penutup. Beban sudah memperhitungkan
    persediaan awal dan akhir jika penyesuaian sudah dibuat."""
    s = saldo_per_akun(df)
    awal, akhir = nilai_persediaan_iktisar(df)
    pendapatan = total_tipe(s, "Pendapatan")
    beban = total_tipe(s, "Beban") + awal - akhir
    return pendapatan, beban, pendapatan - beban


# =====================================================================
# Laporan
# =====================================================================
def laporan_jurnal_umum(df):
    """Jurnal umum murni: tanpa saldo awal. Nama akun saja (tanpa kode),
    akun kredit dijorok ke kanan (aturan penulisan jurnal)."""
    df = df[df["jenis"] != "saldo_awal"]
    j = pd.DataFrame({
        "Tanggal": df["tanggal"],
        "Bukti": df["bukti"].str[:8],
        "Keterangan": df["keterangan"],
        "Akun": [("      " + a) if k > 0 else a for a, k in zip(df["nama_akun"], df["kredit"])],
        "Debit": df["debit"],
        "Kredit": df["kredit"],
    })
    return j.reset_index(drop=True)


def laporan_buku_besar(df, kode):
    """Buku besar satu akun + saldo berjalan."""
    d = df[df["kode_akun"] == kode].sort_values(["tanggal", "id"])
    normal = semua_akun()[kode]["normal"]
    saldo, baris = 0, []
    for _, r in d.iterrows():
        saldo += (r["debit"] - r["kredit"]) if normal == "D" else (r["kredit"] - r["debit"])
        baris.append({"Tanggal": r["tanggal"], "Keterangan": r["keterangan"],
                      "Debit": r["debit"], "Kredit": r["kredit"], "Saldo": saldo})
    return pd.DataFrame(baris, columns=["Tanggal", "Keterangan", "Debit", "Kredit", "Saldo"])


def laporan_neraca_saldo(df):
    """Neraca saldo (trial balance): saldo akhir tiap akun di kolom Debit ATAU Kredit.
    Total Debit harus sama dengan total Kredit."""
    baris = []
    for kode, info in semua_akun().items():
        d = int(df.loc[df["kode_akun"] == kode, "debit"].sum())
        k = int(df.loc[df["kode_akun"] == kode, "kredit"].sum())
        if d == 0 and k == 0:
            continue
        net = d - k
        baris.append({"Kode": kode, "Akun": info["nama"],
                      "Debit": net if net > 0 else 0,
                      "Kredit": -net if net < 0 else 0})
    t = pd.DataFrame(baris, columns=["Kode", "Akun", "Debit", "Kredit"])
    total = pd.DataFrame([{"Kode": "", "Akun": "TOTAL",
                           "Debit": t["Debit"].sum(), "Kredit": t["Kredit"].sum()}])
    return pd.concat([t, total], ignore_index=True)


def laporan_laba_rugi(df):
    """Laporan laba rugi. df = jurnal pada periode (jurnal penutup diabaikan).
    Susunan: Pendapatan - Harga Pokok Penjualan = Laba Kotor - Beban Operasional."""
    df = tanpa_penutup(df)
    akun = semua_akun()
    s = saldo_per_akun(df)
    awal, akhir = nilai_persediaan_iktisar(df)
    baris = []                                    # (keterangan, jumlah)

    baris.append(("PENDAPATAN", None))
    for kode, info in akun.items():
        if info["tipe"] == "Pendapatan" and s[kode] != 0:
            baris.append(("    " + info["nama"], s[kode]))
    pendapatan = total_tipe(s, "Pendapatan")
    baris.append(("Total Pendapatan", pendapatan))
    baris.append(("", None))

    hpp = 0
    akun_hpp = [k for k in HPP_AKUN if k in akun and s[k] != 0]
    if akun_hpp or awal or akhir:
        baris.append(("HARGA POKOK PENJUALAN", None))
        baris.append(("    Persediaan awal", awal))
        pembelian = 0
        for kode in akun_hpp:
            baris.append(("    " + akun[kode]["nama"], s[kode]))
            pembelian += s[kode]
        baris.append(("    Barang tersedia untuk dijual", awal + pembelian))
        baris.append(("    Persediaan akhir", -akhir))
        hpp = awal + pembelian - akhir
        baris.append(("Total Harga Pokok Penjualan", hpp))
        baris.append(("LABA KOTOR", pendapatan - hpp))
        baris.append(("", None))

    baris.append(("BEBAN OPERASIONAL", None))
    total_beban = 0
    for kode, info in akun.items():
        if info["tipe"] == "Beban" and kode not in HPP_AKUN and s[kode] != 0:
            baris.append(("    " + info["nama"], s[kode]))
            total_beban += s[kode]
    baris.append(("Total Beban Operasional", total_beban))
    baris.append(("", None))
    baris.append(("LABA (RUGI) BERSIH", pendapatan - hpp - total_beban))
    return pd.DataFrame(baris, columns=["Keterangan", "Jumlah"])


def laporan_neraca(df):
    """Neraca (laporan posisi keuangan) per tanggal. df = jurnal sampai tanggal itu.
    Return (tabel, total_aset, total_liabilitas_ekuitas)."""
    akun = semua_akun()
    s = saldo_per_akun(df)
    baris = []

    def seksi(judul, tipe):
        baris.append((judul, None))
        total = 0
        for kode, info in akun.items():
            if info["tipe"] != tipe or s[kode] == 0:
                continue
            # aset: sisi debit positif; liabilitas/ekuitas: sisi kredit positif
            sisi_positif = "D" if tipe == "Aset" else "K"
            nilai = s[kode] if info["normal"] == sisi_positif else -s[kode]
            baris.append(("    " + info["nama"], nilai))
            total += nilai
        return total

    aset = seksi("ASET", "Aset")
    baris.append(("TOTAL ASET", aset))
    baris.append(("", None))
    liabilitas = seksi("LIABILITAS", "Liabilitas")
    baris.append(("Total Liabilitas", liabilitas))
    baris.append(("", None))
    ekuitas = seksi("EKUITAS", "Ekuitas")
    # laba yang belum ditutup (jika penutupan belum dibuat) tetap masuk ekuitas
    laba_berjalan = total_tipe(s, "Pendapatan") - total_tipe(s, "Beban")
    if laba_berjalan != 0:
        baris.append(("    Laba (Rugi) Berjalan", laba_berjalan))
        ekuitas += laba_berjalan
    baris.append(("Total Ekuitas", ekuitas))
    baris.append(("", None))
    baris.append(("TOTAL LIABILITAS + EKUITAS", liabilitas + ekuitas))
    return pd.DataFrame(baris, columns=["Keterangan", "Jumlah"]), aset, liabilitas + ekuitas


# =====================================================================
# Penyesuaian & penutupan (metode Iktisar Laba Rugi)
# =====================================================================
def _df_baris(tanggal, jenis, entri):
    """DataFrame sementara dari daftar entri, untuk menghitung saldo sebelum disimpan."""
    return pd.DataFrame(
        [{"tanggal": tanggal, "jenis": jenis, "kode_akun": k, "debit": d, "kredit": c}
         for k, d, c in entri],
        columns=["tanggal", "jenis", "kode_akun", "debit", "kredit"])


def susun_penyesuaian(df, tanggal, persediaan_akhir, perlengkapan_sisa, penyusutan):
    """Susun jurnal penyesuaian + penutupan. df = jurnal sampai tanggal penyesuaian.
    persediaan_akhir = {"Ikan": nilai, "Pakan": ..., "Obat": ..., "Lainnya": ...} (hasil opname)
    Return: daftar (jenis, keterangan, entri), tiap entri = [(kode, debit, kredit), ...]
    dan tiap entri seimbang."""
    s = saldo_per_akun(df)
    daftar = []

    # 1) Persediaan: awal -> Iktisar (D), Persediaan (K).  Akhir -> Persediaan (D), Iktisar (K)
    entri = []
    for kat, kode in AKUN_PERSEDIAAN.items():
        if s[kode] > 0:
            entri += [(IKHTISAR, s[kode], 0), (kode, 0, s[kode])]
    for kat, kode in AKUN_PERSEDIAAN.items():
        akhir = int(persediaan_akhir.get(kat, 0))
        if akhir > 0:
            entri += [(kode, akhir, 0), (IKHTISAR, 0, akhir)]
    if entri:
        daftar.append(("penyesuaian", "Penyesuaian persediaan awal dan akhir", entri))

    # 2) Perlengkapan: saldo buku disesuaikan ke hasil hitung sisa. Selisihnya = beban.
    selisih = int(perlengkapan_sisa) - s[PERLENGKAPAN]
    if selisih < 0:
        daftar.append(("penyesuaian", "Penyesuaian perlengkapan terpakai",
                       [(BEBAN_PERLENGKAPAN, -selisih, 0), (PERLENGKAPAN, 0, -selisih)]))
    elif selisih > 0:
        daftar.append(("penyesuaian", "Penyesuaian perlengkapan (sisa lebih besar dari saldo)",
                       [(PERLENGKAPAN, selisih, 0), (BEBAN_PERLENGKAPAN, 0, selisih)]))

    # 3) Penyusutan peralatan
    if int(penyusutan) > 0:
        daftar.append(("penyesuaian", "Penyusutan peralatan",
                       [(BEBAN_PENYUSUTAN, int(penyusutan), 0), (AKUM_PENYUSUTAN, 0, int(penyusutan))]))

    # saldo setelah semua penyesuaian
    cur = df
    for jenis, _, entri in daftar:
        cur = pd.concat([cur, _df_baris(tanggal, jenis, entri)], ignore_index=True)

    # 4) Penutupan: pendapatan & beban -> Iktisar, Iktisar -> Modal, Prive -> Modal
    akun = semua_akun()
    for tipe, nama in (("Pendapatan", "Penutupan pendapatan ke Iktisar Laba Rugi"),
                       ("Beban", "Penutupan beban ke Iktisar Laba Rugi")):
        s2 = saldo_per_akun(cur)
        entri = []
        for kode, info in akun.items():
            if info["tipe"] != tipe or s2[kode] == 0:
                continue
            net = s2[kode] if info["normal"] == "D" else -s2[kode]     # debit - kredit
            if net > 0:       # saldo di debit -> dikreditkan, lawannya Iktisar di debit
                entri += [(IKHTISAR, net, 0), (kode, 0, net)]
            else:
                entri += [(kode, -net, 0), (IKHTISAR, 0, -net)]
        if entri:
            daftar.append(("penutup", nama, entri))
            cur = pd.concat([cur, _df_baris(tanggal, "penutup", entri)], ignore_index=True)

    s2 = saldo_per_akun(cur)
    laba = s2[IKHTISAR]                          # saldo kredit iktisar = laba
    if laba > 0:
        entri = [(IKHTISAR, laba, 0), (MODAL, 0, laba)]
        daftar.append(("penutup", "Penutupan Iktisar Laba Rugi ke Modal (laba)", entri))
    elif laba < 0:
        entri = [(MODAL, -laba, 0), (IKHTISAR, 0, -laba)]
        daftar.append(("penutup", "Penutupan Iktisar Laba Rugi ke Modal (rugi)", entri))
    if s2[PRIVE] > 0:
        daftar.append(("penutup", "Penutupan Prive ke Modal",
                       [(MODAL, s2[PRIVE], 0), (PRIVE, 0, s2[PRIVE])]))
    return daftar


# =====================================================================
# Penyusunan jurnal dan rincian saldo awal
# =====================================================================
def entri_saldo_awal(d):
    """Susun jurnal saldo awal:
       Debit  : kas, kas di bank, piutang, persediaan, perlengkapan, peralatan
       Kredit : utang bank, utang orang
       Modal  : selisihnya (aset - utang) supaya jurnal seimbang."""
    saldo = {}

    def tambah(kode, jumlah):
        if jumlah:
            saldo[kode] = saldo.get(kode, 0) + jumlah

    total = lambda daftar: sum(r["nominal"] for r in daftar)
    tambah(KAS_TANGAN, d.get("kas_tangan", 0))
    tambah(KAS_BANK, d["kas"])
    tambah("1-102", total(d["piutang"]))
    tambah("1-103", total(d["ikan"]))
    tambah("1-104", total(d["pakan"]))
    tambah("1-105", total(d["obat"]))
    tambah("1-106", total(d["lainnya"]))
    tambah("1-107", total(d["perlengkapan"]))
    tambah("1-108", total(d["peralatan"]))
    for r in d["utang"]:
        tambah("2-101" if r["pihak"] == "Bank" else "2-102", r["nominal"])

    akun = semua_akun()
    entri = []
    for kode, jml in saldo.items():
        if akun[kode]["normal"] == "D":
            entri.append((kode, jml, 0))
        else:
            entri.append((kode, 0, jml))
    selisih = sum(e[1] for e in entri) - sum(e[2] for e in entri)
    if selisih > 0:
        entri.append((MODAL, 0, selisih))
    elif selisih < 0:
        entri.append((MODAL, -selisih, 0))
    return entri


def item_saldo_awal(d):
    """Rincian saldo awal (jenis ikan, harga, umur pemakaian, dll) untuk tabel saldo_awal_item."""
    items = []
    if d.get("kas_tangan"):
        items.append({"kelompok": "kas_tangan", "pihak": None, "nama": "Kas",
                      "nominal": d["kas_tangan"], "umur_bulan": None})
    if d["kas"]:
        items.append({"kelompok": "kas_bank", "pihak": None, "nama": "Kas di bank",
                      "nominal": d["kas"], "umur_bulan": None})
    for kelompok in ("utang", "piutang", "ikan", "pakan", "obat", "lainnya",
                     "peralatan", "perlengkapan"):
        for r in d[kelompok]:
            items.append({"kelompok": kelompok, "pihak": r.get("pihak"), "nama": r["nama"],
                          "nominal": r["nominal"], "umur_bulan": r.get("umur")})
    return items


# ####################################################################
# BAGIAN 3. DATABASE (Supabase)
# (asal: database.py)
# ####################################################################

# ---------- IMPORT (memanggil library yang dipakai) ----------



# ---------- Koneksi Supabase ----------
@st.cache_resource          # dibuat sekali lalu dipakai ulang (tidak membuat koneksi baru tiap klik)
def _buat_client(url, key):
    return create_client(url, key)


def get_client():
    """Ambil koneksi Supabase (pakai service_role key dari secrets.toml)."""
    try:
        return _buat_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    except Exception:
        st.error("Isi SUPABASE_URL dan SUPABASE_SERVICE_KEY di .streamlit/secrets.toml")
        st.stop()                      # hentikan program di sini


def uid():
    """ID akun yang sedang login. Dipakai di SEMUA query supaya data antar akun tidak tercampur."""
    return st.session_state.user["id"]

# =====================================================================
# 3. DATABASE: ambil & simpan jurnal
# =====================================================================
KOLOM_JURNAL = ["id", "bukti", "tanggal", "keterangan", "jenis", "kode_akun", "debit", "kredit"]


def ambil_semua(tabel):
    """Ambil semua baris tabel. Supabase membatasi 1000 baris per permintaan,
    jadi kita ambil bertahap (halaman demi halaman)."""
    hasil, mulai = [], 0
    while True:
        res = get_client().table(tabel).select("*").eq("user_id", uid()).order("id").range(mulai, mulai + 999).execute()
        hasil += res.data
        if len(res.data) < 1000:
            break
        mulai += 1000
    return hasil


def muat_jurnal():
    """Ambil seluruh jurnal user sebagai DataFrame. Disimpan sementara (cache)
    di session_state supaya tidak query berulang tiap klik."""
    if "jurnal_df" not in st.session_state:
        df = pd.DataFrame(ambil_semua("jurnal"), columns=KOLOM_JURNAL)
        df["tanggal"] = pd.to_datetime(df["tanggal"]).dt.date
        df[["debit", "kredit"]] = df[["debit", "kredit"]].astype("int64")
        akun = semua_akun()
        df["nama_akun"] = df["kode_akun"].map(lambda k: akun.get(k, {"nama": k})["nama"])
        df["akun"] = df["kode_akun"].map(lambda k: f"{k} - {akun.get(k, {'nama': k})['nama']}")
        st.session_state.jurnal_df = df.sort_values(["tanggal", "id"]).reset_index(drop=True)
    return st.session_state.jurnal_df


def bersihkan_cache():
    """Hapus cache jurnal -> data diambil ulang dari database di pemuatan berikutnya."""
    st.session_state.pop("jurnal_df", None)


def baris_jurnal(tanggal, keterangan, jenis, entri):
    """Bentuk baris jurnal siap simpan.
    entri = daftar (kode_akun, debit, kredit). Wajib seimbang!"""
    total_d = sum(e[1] for e in entri)
    total_k = sum(e[2] for e in entri)
    if total_d != total_k:
        raise ValueError(f"Jurnal tidak seimbang: debit {total_d} != kredit {total_k}")
    bukti = str(uuid.uuid4())   # satu kode untuk semua baris dalam satu transaksi
    # 'for kode, d, k in entri' = membongkar tiap tuple (kode, debit, kredit)
    return [
        {"user_id": uid(), "bukti": bukti, "tanggal": tanggal.isoformat(), "keterangan": keterangan,
         "jenis": jenis, "kode_akun": kode, "debit": int(d), "kredit": int(k)}
        for kode, d, k in entri
    ]


def simpan_transaksi(tanggal, keterangan, jenis, entri):
    """Simpan satu transaksi (beberapa baris jurnal sekaligus) ke Supabase."""
    get_client().table("jurnal").insert(baris_jurnal(tanggal, keterangan, jenis, entri)).execute()
    bersihkan_cache()


def simpan_banyak(tanggal, daftar):
    """Simpan beberapa jurnal sekaligus dalam SATU insert (semua berhasil atau semua gagal).
    daftar = [(jenis, keterangan, entri), ...]"""
    baris = []
    for jenis, keterangan, entri in daftar:
        baris += baris_jurnal(tanggal, keterangan, jenis, entri)
    if baris:
        get_client().table("jurnal").insert(baris).execute()
    bersihkan_cache()


# =====================================================================
# COA tambahan (tabel coa_custom)
# =====================================================================
def muat_coa():
    """Ambil akun tambahan milik user ke session_state (sekali per sesi)."""
    if "akun_custom" in st.session_state:
        return
    try:
        res = (get_client().table("coa_custom").select("*").eq("user_id", uid()).execute())
        st.session_state.akun_custom = {
            r["kode"]: {"nama": r["nama"], "tipe": r["tipe"], "normal": r["normal"]}
            for r in res.data}
    except Exception as e:
        st.session_state.akun_custom = {}
        st.warning(f"Tabel coa_custom belum bisa dibaca (sudah menjalankan schema_update.sql?): {e}")


def tambah_akun(kode, nama, tipe, normal):
    get_client().table("coa_custom").insert(
        {"user_id": uid(), "kode": kode, "nama": nama, "tipe": tipe, "normal": normal}).execute()
    st.session_state.akun_custom[kode] = {"nama": nama, "tipe": tipe, "normal": normal}
    bersihkan_cache()          # nama akun di jurnal dibaca ulang


def hapus_akun(kode):
    get_client().table("coa_custom").delete().eq("user_id", uid()).eq("kode", kode).execute()
    st.session_state.akun_custom.pop(kode, None)
    bersihkan_cache()


# =====================================================================
# Daftar harga persediaan (tabel daftar_persediaan)
# =====================================================================
KOLOM_HARGA = ["kategori", "nama", "satuan", "harga_satuan"]


def muat_harga():
    """Daftar barang + harga jual satuan sebagai DataFrame."""
    if "harga_df" not in st.session_state:
        res = (get_client().table("daftar_persediaan").select("*")
               .eq("user_id", uid()).order("kategori").order("nama").execute())
        df = pd.DataFrame(res.data, columns=KOLOM_HARGA + ["id"]).drop(columns="id")
        df["harga_satuan"] = df["harga_satuan"].astype("int64")
        st.session_state.harga_df = df
    return st.session_state.harga_df


def simpan_harga(df):
    """Ganti seluruh daftar harga user dengan isi df (kolom: kategori, nama, satuan, harga_satuan)."""
    sb, u = get_client(), uid()
    baris = [{"user_id": u, **{k: r[k] for k in KOLOM_HARGA}} for r in df.to_dict("records")]
    sb.table("daftar_persediaan").delete().eq("user_id", u).execute()
    if baris:
        sb.table("daftar_persediaan").insert(baris).execute()
    st.session_state.pop("harga_df", None)


def muat_item_peralatan():
    """Rincian peralatan dari saldo awal (nama, nominal, umur_bulan) untuk saran penyusutan."""
    res = (get_client().table("saldo_awal_item").select("*")
           .eq("user_id", uid()).eq("kelompok", "peralatan").execute())
    return res.data

# ---------- Status saldo awal ----------
def saldo_awal_selesai():
    """Cek di database: apakah akun ini sudah menyelesaikan saldo awal?"""
    res = (get_client().table("profil").select("saldo_awal_selesai")
           .eq("user_id", uid()).execute())
    return bool(res.data and res.data[0]["saldo_awal_selesai"])

# ---------- Penyimpanan saldo awal ----------
def simpan_saldo_awal(d, tanggal):
    """Simpan saldo awal ke Supabase: jurnal + rincian + tanda 'selesai'."""
    sb = get_client()
    u = uid()
    entri = entri_saldo_awal(d)
    # hapus sisa percobaan sebelumnya milik akun ini (jika sempat gagal di tengah jalan)
    sb.table("jurnal").delete().eq("user_id", u).eq("jenis", "saldo_awal").execute()
    sb.table("saldo_awal_item").delete().eq("user_id", u).execute()
    if entri:
        sb.table("jurnal").insert(
            baris_jurnal(tanggal, "Saldo awal", "saldo_awal", entri)).execute()
    items = [{**it, "user_id": u} for it in item_saldo_awal(d)]   # tambahkan user_id ke tiap baris
    if items:
        sb.table("saldo_awal_item").insert(items).execute()
    # terakhir: tandai selesai (upsert = insert, atau update jika sudah ada)
    sb.table("profil").upsert({"user_id": u, "saldo_awal_selesai": True}).execute()
    bersihkan_cache()
    st.session_state.sa_selesai = True
    st.session_state.pop("sa", None)


# ####################################################################
# BAGIAN 4. FUNGSI BANTU TAMPILAN & DASHBOARD
# (asal: supports.py)
# ####################################################################

# =====================================================================
# 2. FUNGSI BANTU TAMPILAN
# =====================================================================
def rp(angka):
    """Ubah angka jadi teks rupiah: 1500000 -> 'Rp 1.500.000'."""
    angka = int(angka)
    # f"{x:,}" memberi koma ribuan (1,500,000), lalu koma kita ganti jadi titik
    teks = f"{abs(angka):,}".replace(",", ".")
    return f"-Rp {teks}" if angka < 0 else f"Rp {teks}"

def ke_excel(sheets):
    """Buat file Excel (bytes). sheets = {"Nama Sheet": DataFrame, ...}"""
    buf = io.BytesIO()                                   # 'file' sementara di memori
    with pd.ExcelWriter(buf, engine="openpyxl") as w:    # 'with' = otomatis tutup file
        for nama, df in sheets.items():
            nama = nama[:31]                             # batas nama sheet Excel: 31 huruf
            df.to_excel(w, sheet_name=nama, index=False)
            ws = w.sheets[nama]
            # atur lebar tiap kolom sesuai isi
            for i, kol in enumerate(df.columns, start=1):
                isi = [len(str(kol))] + [len(str(x)) for x in df[kol].head(200)]
                ws.column_dimensions[get_column_letter(i)].width = min(max(isi) + 2, 50)
    return buf.getvalue()


def tombol_excel(label, sheets, nama_file, kunci):
    """Tombol unduh Excel. 'kunci' harus unik per tombol."""
    st.download_button(
        label, data=ke_excel(sheets), file_name=nama_file, key=kunci,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def tombol_keluar():
    """Menampilkan tombol keluar di bagian kanan atas halaman."""
    kiri, kanan = st.columns([8, 1])

    with kanan:
        if st.button("Keluar", key="tombol_keluar"):
            st.session_state.clear()
            st.rerun()

def format_uang(df, kolom):
    """Salinan DataFrame dengan kolom uang diformat 'Rp ...' (untuk tampilan saja)."""
    d = df.copy()
    for k in kolom:
        # lambda = fungsi mini. Kosong (NaN) ditampilkan kosong, selain itu diformat rupiah
        d[k] = d[k].map(lambda v: "" if pd.isna(v) else rp(v))
    return d

def judul_halaman(judul, deskripsi):
    """Kepala halaman dengan tombol keluar di kanan atas."""

    tombol_keluar()

    st.title(judul)
    st.caption(deskripsi)
    
# =====================================================================
# 7. HALAMAN DASHBOARD
# =====================================================================
PERIODE = ["Bulan ini", "30 hari", "Tahun ini", "Kustom"]


def pilih_periode(kunci, bawaan):
    """Bar filter periode: tombol cepat (Bulan ini / 30 hari / Tahun ini) atau Kustom.
    'kunci' harus unik per halaman. Return (tanggal_awal, tanggal_akhir)."""
    hari = date.today()
    with st.container(border=True):
        c1, c2, c3 = st.columns([2, 1.5, 1.5], vertical_alignment="bottom")
        pilih = c1.segmented_control("Periode", PERIODE, default=bawaan,
                                     key=f"{kunci}_periode") or bawaan
        if pilih == "Kustom":
            awal = c2.date_input("Dari tanggal", value=hari.replace(day=1), key=f"{kunci}_awal")
            akhir = c3.date_input("Sampai tanggal", value=hari, key=f"{kunci}_akhir")
        else:
            awal = {"Bulan ini": hari.replace(day=1),
                    "30 hari": hari - timedelta(days=29),
                    "Tahun ini": date(hari.year, 1, 1)}[pilih]
            akhir = hari
            c2.caption(f"Dari {awal.strftime('%d/%m/%Y')}")
            c3.caption(f"Sampai {akhir.strftime('%d/%m/%Y')}")
    if awal > akhir:
        st.error("Tanggal awal tidak boleh setelah tanggal akhir.")
        st.stop()
    return awal, akhir


def delta_rp(sekarang, sebelum):
    """Teks selisih untuk st.metric. Negatif -> diawali '-' (otomatis merah). Nol -> tanpa delta."""
    selisih = sekarang - sebelum
    return None if selisih == 0 else rp(selisih)


def data_grafik(df, awal, akhir):
    """Pendapatan vs Beban per hari (periode pendek) atau per bulan (periode panjang).
    Return DataFrame berindeks 'Periode' dengan kolom Pendapatan & Beban, atau None jika kosong."""
    d = filter_periode(df, awal, akhir)
    akun = semua_akun()
    tipe = d["kode_akun"].map(lambda k: akun[k]["tipe"])
    d = d[tipe.isin(["Pendapatan", "Beban"])].copy()
    if d.empty:
        return None
    tipe = d["kode_akun"].map(lambda k: akun[k]["tipe"])
    d["Pendapatan"] = (d["kredit"] - d["debit"]).where(tipe == "Pendapatan", 0)
    d["Beban"] = (d["debit"] - d["kredit"]).where(tipe == "Beban", 0)
    format_waktu = "%Y-%m-%d" if (akhir - awal).days <= 62 else "%Y-%m"
    d["Periode"] = pd.to_datetime(d["tanggal"]).dt.strftime(format_waktu)
    return d.groupby("Periode")[["Pendapatan", "Beban"]].sum().sort_index()


def tabel_transaksi(df, jenis=None):
    """Daftar transaksi pembelian / penjualan (satu baris per transaksi), terbaru di atas.
    jenis=None -> gabungan keduanya (ada kolom 'Jenis')."""
    m_beli = (df["jenis"] == "pembelian") & (df["debit"] > 0)     # sisi beban/aset
    m_jual = (df["jenis"] == "penjualan") & (df["kredit"] > 0)    # sisi pendapatan
    if jenis == "pembelian":
        t = df[m_beli]
    elif jenis == "penjualan":
        t = df[m_jual]
    else:
        t = df[m_beli | m_jual]
    t = t.sort_values(["tanggal", "id"], ascending=False)
    akun = semua_akun()

    def pisah(ket, kode):
        """Keterangan 'Ikan: Nila 10 kg' -> kategori 'Ikan', keterangan 'Nila 10 kg'."""
        for kat in KATEGORI_BARANG:
            if ket.startswith(kat + ": "):
                return kat, ket[len(kat) + 2:]
        return akun.get(kode, {"nama": kode})["nama"], ket

    pasangan = [pisah(k, c) for k, c in zip(t["keterangan"], t["kode_akun"])]
    tabel = pd.DataFrame({
        "Tanggal": t["tanggal"],
        "Jenis": t["jenis"].astype(str).str.capitalize(),
        "Kategori": [p[0] for p in pasangan],
        "Keterangan": [p[1] for p in pasangan],
        "Nominal": t["debit"] + t["kredit"],
    })
    return tabel.drop(columns="Jenis") if jenis else tabel


def halaman_dashboard():
    judul_halaman("Dashboard", f"Ringkasan keuangan  ·  {st.session_state.user['nama']}")
    df = tanpa_penutup(muat_jurnal())      # jurnal penutup tidak ikut, supaya pendapatan/beban tetap terlihat
    awal, akhir = pilih_periode("dash", "Bulan ini")

    # Periode pembanding = periode sebelumnya dengan jumlah hari yang sama
    panjang = (akhir - awal).days + 1
    akhir_lalu = awal - timedelta(days=1)
    awal_lalu = akhir_lalu - timedelta(days=panjang - 1)

    # Laba/pendapatan/beban = hanya dalam periode.  Kas/utang/piutang = saldo sampai tanggal akhir.
    s_akhir = saldo_per_akun(filter_periode(df, None, akhir))
    s_awal = saldo_per_akun(filter_periode(df, None, akhir_lalu))

    pendapatan, beban, laba = hitung_laba(filter_periode(df, awal, akhir))
    pend_lalu, beban_lalu, laba_lalu = hitung_laba(filter_periode(df, awal_lalu, akhir_lalu))
    kas, bank, piutang = s_akhir[KAS_TANGAN], s_akhir[KAS_BANK], s_akhir["1-102"]
    utang_bank, utang_orang = s_akhir["2-101"], s_akhir["2-102"]

    # ---- Baris 1: 4 kartu angka utama ----
    k = st.columns(5)
    banding = "Dibanding periode sebelumnya (jumlah hari sama)"
    k[0].metric("Pendapatan", rp(pendapatan), delta_rp(pendapatan, pend_lalu),
                border=True, help=banding)
    k[1].metric("Beban", rp(beban), delta_rp(beban, beban_lalu), delta_color="inverse",
                border=True, help=banding)
    k[2].metric("Laba (Rugi)", rp(laba), delta_rp(laba, laba_lalu),
                border=True, help=banding)
    k[3].metric("Kas", rp(kas), delta_rp(kas, s_awal[KAS_TANGAN]),
                border=True, help="Perubahan dibanding saldo kas di awal periode")
    k[4].metric("Kas di Bank", rp(bank), delta_rp(bank, s_awal[KAS_BANK]),
                border=True, help="Perubahan dibanding saldo bank di awal periode")

    # ---- Baris 2: grafik (kiri, lebar) + utang/piutang (kanan) ----
    kiri, kanan = st.columns([2, 1], gap="medium")
    with kiri:
        with st.container(border=True):
            st.subheader("Pendapatan vs Beban")
            grafik = data_grafik(df, awal, akhir)
            if grafik is None:
                st.info("Belum ada pendapatan atau beban pada periode ini.")
            else:
                st.bar_chart(grafik, color=["#16a34a", "#dc2626"], stack=False)
    with kanan:
        with st.container(border=True):
            st.subheader("Utang / Piutang")
            st.metric("Piutang", rp(piutang))
            st.metric("Total Utang", rp(utang_bank + utang_orang))
            st.caption(f"Utang Bank: {rp(utang_bank)}  |  Utang Orang: {rp(utang_orang)}")
            st.divider()
            st.metric("Piutang - Utang", rp(piutang - utang_bank - utang_orang))

    # ---- Baris 3: transaksi terbaru ----
    with st.container(border=True):
        st.subheader("Transaksi Terbaru")
        terbaru = tabel_transaksi(df).head(8)
        if terbaru.empty:
            st.info("Belum ada transaksi. Mulai dari menu Pembelian atau Penjualan.")
        else:
            st.dataframe(format_uang(terbaru, ["Nominal"]), hide_index=True)


# ####################################################################
# BAGIAN 5. LOGIN & SIDEBAR AKUN
# (asal: login.py)
# ####################################################################

# =====================================================================
# 5. LOGIN & SIDEBAR AKUN  (akun di tabel `akun`, tanpa Supabase Auth)
# =====================================================================
def halaman_login():
    # Tiga kolom: kiri & kanan kosong, kartu login di tengah -> posisi di tengah layar
    _, tengah, _ = st.columns([1, 1.2, 1])
    with tengah:
        st.write("")
        st.write("")
        with st.container(border=True):
            st.header("🐟 SiCupang")
            st.caption("Aplikasi akuntansi sederhana (cash basis)")
            with st.form("form_masuk"):                   # form = kirim sekali saat tombol ditekan
                username = st.text_input("Username", placeholder="Masukkan username")
                password = st.text_input("Password", type="password", placeholder="Masukkan password")
                kirim = st.form_submit_button("Masuk", type="primary")
        if not kirim:
            return
        if not username.strip() or not password:
            st.error("Username dan password wajib diisi.")
            return
        try:
            # panggil fungsi database masuk_akun(); hasil: [] jika salah, [{id, nama}] jika benar
            res = get_client().rpc("masuk_akun", {"p_username": username.strip(),
                                                  "p_password": password}).execute()
        except Exception as e:
            st.error(f"Gagal terhubung ke database: {e}")
            return
        if not res.data:
            st.error("Username atau password salah.")
            return
        akun = res.data[0]
        st.session_state.user = {"id": akun["id"], "nama": akun["nama"]}
        bersihkan_cache()
        st.rerun()                                        # muat ulang halaman


def keluar():
    """Logout: kosongkan semua memori sesi."""
    st.session_state.clear()
    st.rerun()


def sidebar_akun():
    """Kartu akun di bagian bawah sidebar (di bawah menu): nama, info, tombol keluar."""
    with st.sidebar:
        st.divider()
        with st.container(border=True):
            st.markdown("**🐟 SiCupang**")
            st.caption(f":material/account_circle: {st.session_state.user['nama']}")
            st.caption("Pembukuan cash basis")
            if st.button("Keluar", icon=":material/logout:", key="btn_keluar"):
                keluar()


# ####################################################################
# BAGIAN 6. SALDO AWAL (WIZARD)
# (asal: saldo_awal.py)
# ####################################################################

# =====================================================================
# 6. SALDO AWAL (wizard 7 langkah, pertama kali login)
# =====================================================================
TOTAL_LANGKAH = 7
OPSI = ["Tidak", "Ya"]

# Definisi tabel isian. Tiap kolom: (judul kolom, peran)
#   peran: pihak (Bank/Orang) | nama (teks) | nominal (rupiah) | umur (bulan)
# Ingin ubah judul kolom / tambah kolom? Ubah di sini.
TABEL = {
    "utang":        [("Pihak", "pihak"), ("Nama bank / orang", "nama"), ("Nominal (Rp)", "nominal")],
    "piutang":      [("Nama pihak", "nama"), ("Nominal (Rp)", "nominal")],
    "ikan":         [("Jenis ikan", "nama"), ("Harga (Rp)", "nominal")],
    "pakan":        [("Nama pakan", "nama"), ("Harga (Rp)", "nominal")],
    "obat":         [("Nama obat", "nama"), ("Harga (Rp)", "nominal")],
    "lainnya":      [("Nama barang", "nama"), ("Harga (Rp)", "nominal")],
    "peralatan":    [("Nama peralatan", "nama"), ("Harga (Rp)", "nominal"),
                     ("Perkiraan pemakaian (bulan)", "umur")],
    "perlengkapan": [("Nama perlengkapan", "nama"), ("Harga (Rp)", "nominal")],
}


def init_wizard():
    """Siapkan 'memori' wizard saat pertama dibuka."""
    if "sa" not in st.session_state:
        st.session_state.sa = {
            "langkah": 0,
            "kas_tangan": None,                      # wajib diisi (None = belum diisi)
            "kas_ada": "Tidak", "kas_nominal": 0,
            "utang_ada": "Tidak", "piutang_ada": "Tidak", "lain_ada": "Tidak",
            "tanggal": date.today(),
            "in": {},    # tabel awal yang ditampilkan editor
            "out": {},   # hasil editan terbaru dari editor
        }


def tanya(kunci, pertanyaan):
    """Pertanyaan Ya/Tidak. Jawaban disimpan di sa[kunci]. Return True jika 'Ya'."""
    sa = st.session_state.sa
    sa[kunci] = st.radio(pertanyaan, OPSI, index=OPSI.index(sa[kunci]),
                         horizontal=True, key=f"radio_{kunci}")
    return sa[kunci] == "Ya"


def editor(nama):
    """Tampilkan tabel isian yang bisa ditambah barisnya (klik baris kosong di bawah)."""
    sa = st.session_state.sa
    if nama not in sa["in"]:
        # tabel kosong dengan kolom sesuai TABEL (teks -> object, angka -> float)
        sa["in"][nama] = pd.DataFrame({
            kol: pd.Series(dtype="object" if peran in ("pihak", "nama") else "float64")
            for kol, peran in TABEL[nama]
        })
    config = {}                                       # pengaturan tampilan tiap kolom
    for kol, peran in TABEL[nama]:
        if peran == "pihak":
            config[kol] = st.column_config.SelectboxColumn(kol, options=["Bank", "Orang"])
        elif peran == "nominal":
            config[kol] = st.column_config.NumberColumn(kol, min_value=0, step=1000, format="%d")
        elif peran == "umur":
            config[kol] = st.column_config.NumberColumn(kol, min_value=0, step=1, format="%d")
    # num_rows="dynamic" = pengguna boleh menambah/menghapus baris
    sa["out"][nama] = st.data_editor(sa["in"][nama], num_rows="dynamic",
                                     column_config=config, key=f"editor_{nama}")


def bersih(nama):
    """Ambil isi tabel, buang baris kosong, cek kelengkapan.
    Return: list dict {'pihak','nama','nominal','umur'}. Raise ValueError jika ada yang salah."""
    df = st.session_state.sa["out"].get(nama)
    if df is None:
        return []
    hasil = []
    for no, (_, r) in enumerate(df.iterrows(), start=1):
        isi, kosong = {}, True
        for kol, peran in TABEL[nama]:
            v = r[kol]
            if peran in ("pihak", "nama"):
                v = "" if pd.isna(v) else str(v).strip()
            else:
                v = 0 if pd.isna(v) else int(v)
            if v:
                kosong = False
            isi[peran] = v
        if kosong:
            continue                                  # baris kosong dilewati
        for kol, peran in TABEL[nama]:
            if not isi[peran]:
                raise ValueError(f"Tabel {nama}, baris {no}: kolom '{kol}' belum diisi.")
        hasil.append(isi)
    return hasil


def kumpulkan():
    """Kumpulkan semua jawaban wizard menjadi satu dict data bersih."""
    sa = st.session_state.sa
    return {
        "kas_tangan": int(sa["kas_tangan"] or 0),
        "kas": int(sa["kas_nominal"]) if sa["kas_ada"] == "Ya" else 0,
        "utang": bersih("utang") if sa["utang_ada"] == "Ya" else [],
        "piutang": bersih("piutang") if sa["piutang_ada"] == "Ya" else [],
        "ikan": bersih("ikan"),
        "pakan": bersih("pakan"),
        "obat": bersih("obat"),
        "lainnya": bersih("lainnya") if sa["lain_ada"] == "Ya" else [],
        "peralatan": bersih("peralatan"),
        "perlengkapan": bersih("perlengkapan"),
    }


def validasi_langkah(l):
    """Cek isian langkah ke-l. Return teks error, atau None jika aman."""
    sa = st.session_state.sa
    try:
        if l == 1 and sa["kas_tangan"] is None:
            return "Kas (uang tunai di tangan) wajib diisi. Isi 0 jika tidak ada."
        if l == 1 and sa["kas_ada"] == "Ya" and sa["kas_nominal"] <= 0:
            return "Nominal kas di bank harus lebih dari 0."
        if l == 2 and sa["utang_ada"] == "Ya" and not bersih("utang"):
            return "Isi minimal satu utang, atau pilih 'Tidak'."
        if l == 3 and sa["piutang_ada"] == "Ya" and not bersih("piutang"):
            return "Isi minimal satu piutang, atau pilih 'Tidak'."
        if l == 4:
            for t in ("ikan", "pakan", "obat"):
                bersih(t)
            if sa["lain_ada"] == "Ya" and not bersih("lainnya"):
                return "Isi persediaan lainnya, atau pilih 'Tidak'."
        if l == 5:
            bersih("peralatan")
        if l == 6:
            bersih("perlengkapan")
    except ValueError as e:                           # pesan dari bersih()
        return str(e)
    return None


def pindah(n):
    """Pindah ke langkah n. Hasil editan disalin supaya tidak hilang saat kembali."""
    sa = st.session_state.sa
    sa["in"] = {k: v.copy() for k, v in sa["out"].items()}
    sa["langkah"] = n
    st.rerun()


NAMA_LANGKAH = ["Mulai", "Kas dan Bank", "Utang", "Piutang",
                "Persediaan", "Peralatan", "Perlengkapan"]


def halaman_saldo_awal():
    init_wizard()
    sa = st.session_state.sa
    l = sa["langkah"]

    # Kolom kiri & kanan kosong -> isi wizard berada di tengah, tidak melebar
    _, tengah, _ = st.columns([1, 3, 1])
    with tengah:
        st.title("Saldo Awal")
        st.progress((l + 1) / TOTAL_LANGKAH,
                    text=f"Langkah {l + 1} dari {TOTAL_LANGKAH}  ·  {NAMA_LANGKAH[l]}")

        with st.container(border=True):
            if l == 0:    # 1) pernyataan
                st.subheader("Anda harus memasukkan beberapa saldo untuk memulai❗")
                st.write("Semua isian berupa nominal uang, dipakai sebagai saldo awal pembukuan Anda.")

            elif l == 1:  # 2) kas (tangan) + kas di bank
                st.subheader("Kas dan Bank")
                sa["kas_tangan"] = st.number_input(
                    "Kas: uang tunai di tangan (Rp) — wajib diisi", min_value=0, step=100000,
                    value=sa["kas_tangan"], placeholder="Isi 0 jika tidak ada", key="num_kas_tangan")
                st.divider()
                if tanya("kas_ada", "Apakah Anda memiliki kas di bank?"):
                    sa["kas_nominal"] = st.number_input(
                        "Nominal kas di bank (Rp)", min_value=0, step=100000,
                        value=int(sa["kas_nominal"]), key="num_kas")

            elif l == 2:  # 3) utang
                st.subheader("Utang")
                if tanya("utang_ada", "Apakah Anda memiliki utang?"):
                    st.caption("Pilih Bank atau Orang, isi nama dan nominal. Klik baris kosong di bawah tabel untuk menambah.")
                    editor("utang")

            elif l == 3:  # 4) piutang
                st.subheader("Piutang")
                if tanya("piutang_ada", "Apakah Anda memiliki piutang?"):
                    st.caption("Klik baris kosong di bawah tabel untuk menambah.")
                    editor("piutang")

            elif l == 4:  # 5) persediaan
                st.subheader("Persediaan")
                st.caption("Kosongkan tabel jika tidak ada. Klik baris kosong di bawah tabel untuk menambah.")
                ada_lain = tanya("lain_ada", "Apakah ada persediaan lain selain ikan, pakan, dan obat?")
                st.markdown("**Ikan**")
                editor("ikan")
                st.markdown("**Pakan**")
                editor("pakan")
                st.markdown("**Obat**")
                editor("obat")
                if ada_lain:
                    st.markdown("**Persediaan lainnya**")
                    editor("lainnya")

            elif l == 5:  # 6) peralatan
                st.subheader("Peralatan")
                st.caption("Perkiraan pemakaian = berapa bulan alat diperkirakan bisa dipakai. Kosongkan jika tidak ada.")
                editor("peralatan")

            elif l == 6:  # 7) perlengkapan + ringkasan
                st.subheader("Perlengkapan")
                st.caption("Kosongkan jika tidak ada.")
                editor("perlengkapan")
                sa["tanggal"] = st.date_input("Tanggal saldo awal", value=sa["tanggal"])
                try:                                   # ringkasan sementara
                    entri = entri_saldo_awal(kumpulkan())
                    aset = sum(e[1] for e in entri if e[0].startswith("1-"))
                    utang = sum(e[2] for e in entri if e[0].startswith("2-"))
                    modal = sum(e[2] - e[1] for e in entri if e[0] == "3-101")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Total Aset", rp(aset), border=True)
                    c2.metric("Total Utang", rp(utang), border=True)
                    c3.metric("Modal", rp(modal), border=True)
                except ValueError as e:
                    st.warning(str(e))

        # ---- tombol navigasi ----
        kiri, kanan = st.columns(2)
        kembali = kiri.button("Kembali", icon=":material/arrow_back:", disabled=(l == 0))
        teks = "Oke" if l == 0 else ("Simpan" if l == TOTAL_LANGKAH - 1 else "Lanjut")
        lanjut = kanan.button(teks, type="primary")

        if kembali:
            pindah(l - 1)
        if lanjut:
            err = validasi_langkah(l)
            if err:
                st.error(err)
            elif l == TOTAL_LANGKAH - 1:
                try:
                    simpan_saldo_awal(kumpulkan(), sa["tanggal"])
                except Exception as e:
                    st.error(f"Gagal menyimpan: {e}")
                else:
                    st.rerun()                         # masuk ke aplikasi utama
            else:
                pindah(l + 1)


# ####################################################################
# BAGIAN 7. TRANSAKSI: PEMBELIAN, PENJUALAN, TARIK UANG
# (asal: transaksi.py)
# ####################################################################

NAMA_KAS = {kode: nama for nama, kode in KAS_PILIHAN.items()}


def ringkasan_transaksi(judul, jenis):
    """3 kartu angka di atas halaman: total bulan ini, jumlah transaksi, saldo kas & bank."""
    df = muat_jurnal()
    tabel = tabel_transaksi(df, jenis)
    bulan_ini = tabel[tabel["Tanggal"] >= date.today().replace(day=1)]
    s = saldo_per_akun(df)
    m = st.columns(4)
    m[0].metric(f"Total {judul} bulan ini", rp(bulan_ini["Nominal"].sum()), border=True)
    m[1].metric("Jumlah transaksi bulan ini", f"{len(bulan_ini)}", border=True)
    m[2].metric("Kas", rp(s[KAS_TANGAN]), border=True)
    m[3].metric("Kas di Bank", rp(s[KAS_BANK]), border=True)


def riwayat(judul, jenis):
    """10 transaksi terbaru + tombol unduh Excel."""
    with st.container(border=True):
        st.subheader(f"10 {judul} Terbaru")
        tabel = tabel_transaksi(muat_jurnal(), jenis)
        if tabel.empty:
            st.info("Belum ada data.")
        else:
            st.dataframe(format_uang(tabel.head(10), ["Nominal"]), hide_index=True)
            tombol_excel("⬇️ Unduh semua ke Excel", {judul: tabel},
                         f"{jenis}.xlsx", f"xl_{jenis}")


# =====================================================================
# PEMBELIAN
#   Ikan / Pakan / Obat -> Debit Pembelian, Kredit Kas (atau Kas di Bank)
#   Lainnya (beban, perlengkapan, peralatan) -> Debit akunnya, Kredit Kas
# =====================================================================
def halaman_pembelian():
    judul_halaman("Pembelian", "Catat pembelian baru di kiri, riwayat terbaru ada di kanan.")
    ringkasan = st.container()       # diisi paling akhir supaya sudah memuat transaksi baru
    kiri, kanan = st.columns([2, 3], gap="large")

    with kiri:
        with st.container(border=True):
            st.subheader("Input Pembelian")
            pilihan = pilihan_pembelian()      # beban dari menu Tambah COA otomatis ikut
            with st.form("form_pembelian", clear_on_submit=True):
                c1, c2 = st.columns(2)
                tgl = c1.date_input("Tanggal", value=date.today())
                kategori = c2.selectbox("Kategori", list(pilihan.keys()))
                kas_label = st.selectbox("Dibayar dari", list(KAS_PILIHAN.keys()))
                ket = st.text_input("Keterangan (nama barang / penjual)")
                nominal = st.number_input("Nominal (Rp)", min_value=0, step=1000, value=0)
                kirim = st.form_submit_button("Submit", type="primary")

            if kirim:
                kode, kas = pilihan[kategori], KAS_PILIHAN[kas_label]
                saldo_kas = saldo_per_akun(muat_jurnal())[kas]
                if nominal <= 0:
                    st.error("Nominal harus lebih dari 0.")
                elif not ket.strip():
                    st.error("Keterangan wajib diisi.")
                elif nominal > saldo_kas:
                    st.error(f"Saldo {kas_label} tidak cukup ({rp(saldo_kas)}).")
                else:
                    # Ikan/Pakan/Obat: kategori disimpan di depan keterangan agar tetap terbaca
                    teks = f"{kategori}: {ket.strip()}" if kategori in KATEGORI_BARANG else ket.strip()
                    try:
                        simpan_transaksi(tgl, teks, "pembelian",
                                         [(kode, nominal, 0), (kas, 0, nominal)])
                    except Exception as e:
                        st.error(f"Gagal menyimpan: {e}")
                    else:
                        st.success(f"Tersimpan: {kategori} {rp(nominal)}")
    with kanan:
        riwayat("Pembelian", "pembelian")
    with ringkasan:
        ringkasan_transaksi("Pembelian", "pembelian")


# =====================================================================
# PENJUALAN
#   Pilih barang dari Daftar Harga, isi jumlah -> total = jumlah x harga satuan
#   Debit Kas (atau Kas di Bank), Kredit Penjualan
# =====================================================================
def halaman_penjualan():
    judul_halaman("Penjualan", "Pilih barang, isi jumlah. Harga mengikuti menu Daftar Harga.")
    ringkasan = st.container()
    kiri, kanan = st.columns([2, 3], gap="large")
    # 'ver' dinaikkan setelah submit -> semua isian kembali kosong
    ver = st.session_state.setdefault("ver_jual", 0)

    with kiri:
        with st.container(border=True):
            st.subheader("Input Penjualan")
            if "flash_jual" in st.session_state:
                st.success(st.session_state.pop("flash_jual"))
            harga = muat_harga()

            c1, c2 = st.columns(2)
            tgl = c1.date_input("Tanggal", value=date.today(), key=f"jual_tgl_{ver}")
            kategori = c2.selectbox("Kategori", KATEGORI_BARANG, key=f"jual_kat_{ver}")
            barang = harga[harga["kategori"] == kategori].reset_index(drop=True)

            if barang.empty:
                st.info(f"Belum ada barang kategori {kategori}. Isi dulu di menu Daftar Harga.")
            else:
                idx = st.selectbox(
                    "Barang", range(len(barang)), key=f"jual_brg_{ver}",
                    format_func=lambda i: f"{barang.at[i, 'nama']}  ({rp(barang.at[i, 'harga_satuan'])} / {barang.at[i, 'satuan']})")
                b = barang.loc[idx]
                jumlah = st.number_input(f"Jumlah ({b['satuan']})", min_value=0.0, step=1.0,
                                         value=0.0, format="%g", key=f"jual_jml_{ver}")
                kas_label = st.selectbox("Diterima di", list(KAS_PILIHAN.keys()), key=f"jual_kas_{ver}")
                ket = st.text_input("Pembeli (opsional)", key=f"jual_ket_{ver}")
                total = int(round(jumlah * int(b["harga_satuan"])))
                st.metric("Total", rp(total), border=True,
                          help=f"{jumlah:g} x {rp(b['harga_satuan'])}")

                if st.button("Submit", type="primary", key=f"jual_kirim_{ver}"):
                    if total <= 0:
                        st.error("Jumlah harus lebih dari 0.")
                    else:
                        teks = f"{kategori}: {b['nama']} {jumlah:g} {b['satuan']} x {rp(b['harga_satuan'])}"
                        if ket.strip():
                            teks += f" - {ket.strip()}"
                        try:
                            simpan_transaksi(tgl, teks, "penjualan",
                                             [(KAS_PILIHAN[kas_label], total, 0), (PENJUALAN, 0, total)])
                        except Exception as e:
                            st.error(f"Gagal menyimpan: {e}")
                        else:
                            st.session_state.ver_jual = ver + 1
                            st.session_state.flash_jual = f"Tersimpan: {b['nama']} {rp(total)}"
                            st.rerun()
    with kanan:
        riwayat("Penjualan", "penjualan")
    with ringkasan:
        ringkasan_transaksi("Penjualan", "penjualan")


# =====================================================================
# TARIK UANG (prive)  ->  Debit Prive, Kredit Kas / Kas di Bank
# =====================================================================
def halaman_tarik_uang():
    judul_halaman("Tarik Uang", "Pengambilan uang untuk keperluan pribadi (prive). Mengurangi kas atau bank yang dipilih.")
    df = muat_jurnal()
    s = saldo_per_akun(df)
    if "flash_prive" in st.session_state:
        st.success(st.session_state.pop("flash_prive"))
    m = st.columns(2)
    m[0].metric("Kas", rp(s[KAS_TANGAN]), border=True)
    m[1].metric("Kas di Bank", rp(s[KAS_BANK]), border=True)

    kiri, kanan = st.columns([2, 3], gap="large")
    with kiri:
        with st.container(border=True):
            st.subheader("Tarik Uang")
            with st.form("form_prive", clear_on_submit=True):
                sumber_label = st.radio("Ambil dari mana?", list(KAS_PILIHAN.keys()), horizontal=True)
                c1, c2 = st.columns(2)
                tgl = c1.date_input("Tanggal", value=date.today())
                nominal = c2.number_input("Nominal (Rp)", min_value=0, step=10000, value=0)
                ket = st.text_input("Keterangan (opsional)")
                kirim = st.form_submit_button("Submit", type="primary")
            if kirim:
                sumber = KAS_PILIHAN[sumber_label]
                if nominal <= 0:
                    st.error("Nominal harus lebih dari 0.")
                elif nominal > s[sumber]:
                    st.error(f"Saldo {sumber_label} tidak cukup ({rp(s[sumber])}).")
                else:
                    teks = f"Prive dari {sumber_label}" + (f" - {ket.strip()}" if ket.strip() else "")
                    try:
                        simpan_transaksi(tgl, teks, "prive", [(PRIVE, nominal, 0), (sumber, 0, nominal)])
                    except Exception as e:
                        st.error(f"Gagal menyimpan: {e}")
                    else:
                        st.session_state.flash_prive = f"Tersimpan: tarik {rp(nominal)} dari {sumber_label}"
                        st.rerun()
    with kanan:
        with st.container(border=True):
            st.subheader("Riwayat Penarikan")
            p = df[(df["jenis"] == "prive") & (df["debit"] > 0)].sort_values(["tanggal", "id"], ascending=False)
            if p.empty:
                st.info("Belum ada penarikan.")
            else:
                akun_kas = df[(df["jenis"] == "prive") & (df["kredit"] > 0)].set_index("bukti")["kode_akun"]
                tabel = pd.DataFrame({
                    "Tanggal": p["tanggal"],
                    "Sumber": p["bukti"].map(lambda b: NAMA_KAS.get(akun_kas.get(b), "")),
                    "Keterangan": p["keterangan"],
                    "Nominal": p["debit"],
                })
                st.dataframe(format_uang(tabel.head(15), ["Nominal"]), hide_index=True)
                tombol_excel("⬇️ Unduh ke Excel", {"Tarik Uang": tabel}, "tarik_uang.xlsx", "xl_prive")


# ####################################################################
# BAGIAN 8. AKSES CEPAT
# (asal: akses_cepat.py)
# ####################################################################

def halaman_akses_cepat():
    judul_halaman("Akses Cepat", "Jurnal umum, buku besar, dan neraca saldo. Laba rugi dan neraca ada di menu sendiri.")
    df = muat_jurnal()

    # Filter tanggal untuk semua laporan
    awal, akhir = pilih_periode("aq", "Tahun ini")
    df_p = filter_periode(df, awal, akhir)        # data dalam periode
    df_sd = filter_periode(df, None, akhir)       # data sampai tanggal akhir

    # Siapkan semua laporan (juga dipakai untuk Excel)
    jurnal = laporan_jurnal_umum(df_p)
    laba_rugi = laporan_laba_rugi(df_p)
    neraca_saldo = laporan_neraca_saldo(df_sd)
    neraca = laporan_neraca(df_sd)[0]
    akun = semua_akun()
    buku_besar_semua = pd.concat(
        [laporan_buku_besar(df_sd, k).assign(Akun=f"{k} - {akun[k]['nama']}") for k in akun],
        ignore_index=True)

    # Ringkasan singkat periode
    laba = hitung_laba(tanpa_penutup(df_p))[2]
    s_sd = saldo_per_akun(df_sd)
    m = st.columns(4)
    m[0].metric("Laba (Rugi) periode", rp(laba), border=True)
    m[1].metric("Jumlah transaksi", f"{df_p[df_p['jenis'] != 'saldo_awal']['bukti'].nunique()}", border=True)
    m[2].metric("Kas", rp(s_sd[KAS_TANGAN]), border=True)
    m[3].metric("Kas di Bank", rp(s_sd[KAS_BANK]), border=True)

    # Unduh semua laporan
    with st.container(border=True):
        teks, tombol = st.columns([3, 1], vertical_alignment="center")
        teks.markdown("**Unduh laporan lengkap**")
        teks.caption("Jurnal umum, buku besar, neraca saldo, laba rugi, dan neraca dalam satu file Excel.")
        with tombol:
            tombol_excel("⬇️ Unduh SEMUA laporan ke Excel",
                         {"Jurnal Umum": jurnal, "Buku Besar": buku_besar_semua,
                          "Neraca Saldo": neraca_saldo, "Laba Rugi": laba_rugi, "Neraca": neraca},
                         "laporan_akuntansi.xlsx", "xl_semua")

    tab1, tab2, tab4 = st.tabs(["Jurnal Umum", "Buku Besar", "Neraca Saldo"])

    with tab1:
        st.dataframe(format_uang(jurnal, ["Debit", "Kredit"]), hide_index=True)
        tombol_excel("⬇️ Excel", {"Jurnal Umum": jurnal}, "jurnal_umum.xlsx", "xl_jurnal")

    with tab2:
        # format_func = cara menampilkan pilihan (nilai aslinya tetap kode akun)
        kode = st.selectbox("Pilih akun", list(akun), key="bb_akun",
                            format_func=lambda k: f"{k} - {akun[k]['nama']}")
        bb = laporan_buku_besar(df_sd, kode)
        st.caption(f"Saldo normal: {'Debit' if akun[kode]['normal'] == 'D' else 'Kredit'}")
        st.dataframe(format_uang(bb, ["Debit", "Kredit", "Saldo"]), hide_index=True)
        tombol_excel("⬇️ Excel", {akun[kode]["nama"]: bb}, f"buku_besar_{kode}.xlsx", "xl_bb")

    with tab4:
        st.dataframe(format_uang(neraca_saldo, ["Debit", "Kredit"]), hide_index=True)
        total = neraca_saldo.iloc[-1]                  # baris terakhir = TOTAL
        if total["Debit"] == total["Kredit"]:
            st.success("Seimbang ✅ (total Debit = total Kredit)")
        else:
            st.error("Tidak seimbang ❌ — periksa jurnal")
        tombol_excel("⬇️ Excel", {"Neraca Saldo": neraca_saldo}, "neraca_saldo.xlsx", "xl_ns")


# ####################################################################
# BAGIAN 9. DAFTAR HARGA
# (asal: daftar_harga.py)
# ####################################################################

def halaman_daftar_harga():
    judul_halaman("Daftar Harga", "Harga jual satuan tiap barang. Di halaman Penjualan Anda tinggal memilih barang dan mengisi jumlah.")
    harga = muat_harga()

    with st.container(border=True):
        st.caption("Klik baris kosong di bawah tabel untuk menambah barang. Pilih baris lalu tekan Delete untuk menghapus. "
                   "Perubahan baru tersimpan setelah menekan tombol Simpan.")
        edit = st.data_editor(
            harga, num_rows="dynamic", hide_index=True, key="editor_harga",
            column_config={
                "kategori": st.column_config.SelectboxColumn("Kategori", options=KATEGORI_BARANG, required=True),
                "nama": st.column_config.TextColumn("Nama barang", required=True),
                "satuan": st.column_config.TextColumn("Satuan (kg, ekor, sak, ...)", required=True),
                "harga_satuan": st.column_config.NumberColumn("Harga satuan (Rp)", min_value=1, step=500,
                                                              format="%d", required=True),
            })
        if st.button("Simpan daftar harga", type="primary"):
            d = edit.dropna(how="all").copy()
            d["nama"] = d["nama"].fillna("").astype(str).str.strip()
            d["satuan"] = d["satuan"].fillna("").astype(str).str.strip()
            if d[["kategori", "nama", "satuan", "harga_satuan"]].isna().any().any() \
                    or (d["nama"] == "").any() or (d["satuan"] == "").any():
                st.error("Setiap baris wajib diisi lengkap: kategori, nama, satuan, dan harga.")
            elif d.duplicated(["kategori", "nama"]).any():
                st.error("Ada barang dengan kategori dan nama yang sama.")
            else:
                d["harga_satuan"] = d["harga_satuan"].astype("int64")
                try:
                    simpan_harga(d)
                except Exception as e:
                    st.error(f"Gagal menyimpan: {e}")
                else:
                    st.success("Daftar harga tersimpan.")
                    st.rerun()


# ####################################################################
# BAGIAN 10. TAMBAH COA
# (asal: kelola_coa.py)
# ####################################################################

def saran_kode(tipe):
    """Kode berikutnya untuk tipe akun, mis. Beban -> '5-108'."""
    awal = AWAL_KODE[tipe]
    angka = []
    for kode in semua_akun():
        depan, _, belakang = kode.partition("-")
        if depan == awal and belakang.isdigit():
            angka.append(int(belakang))
    return f"{awal}-{(max(angka) if angka else 100) + 1}"


def halaman_coa():
    judul_halaman("Tambah COA", "Tambah akun baru. Akun bertipe Beban otomatis muncul di pilihan Kategori pada halaman Pembelian.")
    kiri, kanan = st.columns([2, 3], gap="large")

    with kiri:
        with st.container(border=True):
            st.subheader("Akun baru")
            tipe = st.selectbox("Tipe akun", TIPE_AKUN, index=TIPE_AKUN.index("Beban"), key="coa_tipe")
            kode = st.text_input("Kode akun", value=saran_kode(tipe), key=f"coa_kode_{tipe}")
            nama = st.text_input("Nama akun", placeholder="mis. Beban Listrik", key="coa_nama")
            normal = st.radio("Saldo normal", ["D", "K"], horizontal=True, key=f"coa_normal_{tipe}",
                              index=["D", "K"].index(NORMAL_BAWAAN[tipe]),
                              format_func=lambda x: "Debit" if x == "D" else "Kredit",
                              help="Biasanya otomatis sesuai tipe. Ubah hanya untuk akun lawan (kontra).")
            if st.button("Tambah akun", type="primary"):
                akun = semua_akun()
                kode, nama = kode.strip(), nama.strip()
                if not kode or not nama:
                    st.error("Kode dan nama akun wajib diisi.")
                elif kode in akun:
                    st.error(f"Kode {kode} sudah dipakai oleh '{akun[kode]['nama']}'.")
                elif nama.lower() in (a["nama"].lower() for a in akun.values()):
                    st.error("Nama akun sudah ada.")
                else:
                    try:
                        tambah_akun(kode, nama, tipe, normal)
                    except Exception as e:
                        st.error(f"Gagal menyimpan: {e}")
                    else:
                        st.success(f"Akun {kode} - {nama} ditambahkan.")
                        st.rerun()

        custom = st.session_state.get("akun_custom", {})
        if custom:
            with st.container(border=True):
                st.subheader("Hapus akun tambahan")
                dipakai = set(muat_jurnal()["kode_akun"])
                bisa = [k for k in custom if k not in dipakai]
                if not bisa:
                    st.caption("Semua akun tambahan sudah punya transaksi, jadi tidak bisa dihapus.")
                else:
                    pilih = st.selectbox("Akun (hanya yang belum punya transaksi)", bisa,
                                         format_func=lambda k: f"{k} - {custom[k]['nama']}")
                    if st.button("Hapus akun"):
                        try:
                            hapus_akun(pilih)
                        except Exception as e:
                            st.error(f"Gagal menghapus: {e}")
                        else:
                            st.rerun()

    with kanan:
        with st.container(border=True):
            st.subheader("Daftar akun")
            tabel = pd.DataFrame([
                {"Kode": k, "Nama": a["nama"], "Tipe": a["tipe"],
                 "Saldo normal": "Debit" if a["normal"] == "D" else "Kredit",
                 "Sumber": "Bawaan" if k in AKUN_DASAR else "Tambahan"}
                for k, a in semua_akun().items()])
            st.dataframe(tabel, hide_index=True, height=520)


# ####################################################################
# BAGIAN 11. LAPORAN LABA RUGI & NERACA
# (asal: laporan.py)
# ####################################################################

def halaman_laba_rugi():
    judul_halaman("Laporan Laba Rugi", "Pendapatan, harga pokok penjualan, dan beban pada periode yang dipilih.")
    df = muat_jurnal()
    awal, akhir = pilih_periode("lr", "Tahun ini")
    d = filter_periode(df, awal, akhir)
    lr = laporan_laba_rugi(d)

    pend = lr.loc[lr["Keterangan"] == "Total Pendapatan", "Jumlah"].iloc[0]
    laba = lr.loc[lr["Keterangan"] == "LABA (RUGI) BERSIH", "Jumlah"].iloc[0]
    m = st.columns(3)
    m[0].metric("Pendapatan", rp(pend), border=True)
    m[1].metric("Beban + HPP", rp(pend - laba), border=True)
    m[2].metric("Laba (Rugi) Bersih", rp(laba), border=True)

    if not (d["jenis"] == "penyesuaian").any():
        st.info("Belum ada penyesuaian pada periode ini. Persediaan akhir belum dihitung, "
                "jadi laba bisa terlalu kecil. Lakukan di menu Penyesuaian.")
    with st.container(border=True):
        st.dataframe(format_uang(lr, ["Jumlah"]), hide_index=True, height=620)
        tombol_excel("⬇️ Excel", {"Laba Rugi": lr}, "laporan_laba_rugi.xlsx", "xl_lap_lr")


def halaman_neraca():
    judul_halaman("Laporan Neraca", "Posisi aset, liabilitas, dan ekuitas pada tanggal yang dipilih.")
    df = muat_jurnal()
    with st.container(border=True):
        tgl = st.date_input("Per tanggal", value=date.today(), key="neraca_tgl")
    neraca, aset, le = laporan_neraca(filter_periode(df, None, tgl))

    m = st.columns(2)
    m[0].metric("Total Aset", rp(aset), border=True)
    m[1].metric("Total Liabilitas + Ekuitas", rp(le), border=True)
    with st.container(border=True):
        st.dataframe(format_uang(neraca, ["Jumlah"]), hide_index=True, height=620)
        if aset == le:
            st.success("Seimbang ✅ (Aset = Liabilitas + Ekuitas)")
        else:
            st.error(f"Tidak seimbang ❌ — selisih {rp(aset - le)}")
        tombol_excel("⬇️ Excel", {"Neraca": neraca}, "laporan_neraca.xlsx", "xl_lap_neraca")


# ####################################################################
# BAGIAN 12. PENYESUAIAN & PENUTUPAN
# (asal: penyesuaian.py)
# ####################################################################

def halaman_penyesuaian():
    judul_halaman("Penyesuaian", "Stock opname, perlengkapan, penyusutan, lalu penutupan buku lewat Iktisar Laba Rugi.")
    df = muat_jurnal()
    if "flash_penyesuaian" in st.session_state:
        st.success(st.session_state.pop("flash_penyesuaian"))

    with st.container(border=True):
        tgl = st.date_input("Tanggal penyesuaian (akhir periode)", value=date.today(), key="adj_tgl")
        tutup_terakhir = df.loc[df["jenis"] == "penutup", "tanggal"].max() if (df["jenis"] == "penutup").any() else None
        if tutup_terakhir is not None:
            st.caption(f"Penutupan terakhir: {tutup_terakhir.strftime('%d/%m/%Y')}. Periode baru dimulai setelah tanggal itu.")
    d = filter_periode(df, None, tgl)
    s = saldo_per_akun(d)

    # ---------- 1. Persediaan akhir (stock opname) ----------
    with st.container(border=True):
        st.subheader("1. Stock opname: persediaan akhir")
        st.caption("Hitung stok fisik di akhir periode, lalu isi nilainya dalam rupiah (sebesar harga perolehan). "
                   "Persediaan awal diambil otomatis dari saldo buku.")
        akhir = {}
        cols = st.columns(len(AKUN_PERSEDIAAN))
        for c, (kat, kode) in zip(cols, AKUN_PERSEDIAAN.items()):
            c.metric(f"{kat}: awal", rp(s[kode]))
            akhir[kat] = c.number_input(f"{kat}: akhir (Rp)", min_value=0, step=10000, value=0,
                                        key=f"adj_akhir_{kat}")
            if akhir[kat] == 0 and s[kode] > 0:
                c.caption("⚠️ Akhir 0, pastikan stok memang habis.")

    # ---------- 2. Perlengkapan ----------
    with st.container(border=True):
        st.subheader("2. Perlengkapan")
        c1, c2 = st.columns(2)
        c1.metric("Saldo perlengkapan di buku", rp(s[PERLENGKAPAN]))
        sisa = c2.number_input("Perlengkapan yang masih tersisa (Rp)", min_value=0, step=10000, value=0,
                               key="adj_sisa",
                               help="Pembelian perlengkapan dicatat sebagai Beban Perlengkapan. Yang masih tersisa "
                                    "dikembalikan ke aset, sisanya tetap menjadi beban.")

    # ---------- 3. Penyusutan peralatan ----------
    with st.container(border=True):
        st.subheader("3. Penyusutan peralatan")
        try:
            item = muat_item_peralatan()
        except Exception:
            item = []
        per_bulan = sum(r["nominal"] / r["umur_bulan"] for r in item if r.get("umur_bulan"))
        c1, c2, c3 = st.columns(3)
        c1.metric("Saldo peralatan", rp(s[PERALATAN]))
        bulan = c2.number_input("Jumlah bulan yang disusutkan", min_value=0, step=1, value=1, key="adj_bulan")
        saran = int(round(per_bulan * bulan))
        susut = c3.number_input("Beban penyusutan (Rp)", min_value=0, step=1000, value=saran,
                                key=f"adj_susut_{bulan}_{saran}",
                                help="Saran dihitung dari harga dan perkiraan pemakaian peralatan di saldo awal "
                                     "(garis lurus). Peralatan yang dibeli kemudian tidak punya umur pemakaian, "
                                     "tambahkan manual ke angka ini.")
        st.caption(f"Saran otomatis: {rp(per_bulan)} per bulan dari peralatan di saldo awal.")

    # ---------- Pratinjau ----------
    daftar = susun_penyesuaian(d, tgl, akhir, sisa, susut)
    akun = semua_akun()
    baris = []
    for jenis, ket, entri in daftar:
        for kode, deb, kre in entri:
            baris.append({"Jenis": jenis.capitalize(), "Keterangan": ket,
                          "Akun": ("      " if kre > 0 else "") + akun[kode]["nama"],
                          "Debit": deb, "Kredit": kre})
    with st.container(border=True):
        st.subheader("Pratinjau jurnal")
        if not baris:
            st.info("Belum ada jurnal yang akan dibuat.")
        else:
            st.dataframe(format_uang(pd.DataFrame(baris), ["Debit", "Kredit"]), hide_index=True)
            # laba setelah penyesuaian = laba dari jurnal sebelum penutup + jurnal penyesuaian
            pra = pd.concat(
                [tanpa_penutup(d)] + [pd.DataFrame(
                    [{"tanggal": tgl, "jenis": j, "kode_akun": k, "debit": a, "kredit": b}
                     for k, a, b in e]) for j, _, e in daftar if j == "penyesuaian"],
                ignore_index=True)
            pend, beban, laba = hitung_laba(pra)
            m = st.columns(3)
            m[0].metric("Pendapatan", rp(pend), border=True)
            m[1].metric("Beban + HPP", rp(beban), border=True)
            m[2].metric("Laba (Rugi) ditutup ke Modal", rp(laba), border=True)

    # ---------- Simpan ----------
    if st.button("Simpan penyesuaian dan penutupan", type="primary", disabled=not baris):
        if tutup_terakhir is not None and tgl <= tutup_terakhir:
            st.error("Tanggal harus setelah penutupan terakhir.")
        else:
            try:
                simpan_banyak(tgl, daftar)
            except Exception as e:
                st.error(f"Gagal menyimpan: {e}")
            else:
                for k in [k for k in st.session_state if str(k).startswith("adj_")]:
                    st.session_state.pop(k)
                st.session_state.flash_penyesuaian = ("Penyesuaian dan penutupan tersimpan. "
                                                      "Lihat hasilnya di Laporan Laba Rugi dan Laporan Neraca.")
                st.rerun()


# ####################################################################
# BAGIAN 13. APLIKASI UTAMA (asal: app.py)
# ####################################################################

def main():
    # 1) Belum login -> tampilkan halaman login (tanpa sidebar)
    if "user" not in st.session_state:
        halaman_login()
        return

    # 2) Sudah login, tapi saldo awal belum diisi -> wizard
    if "sa_selesai" not in st.session_state:
        st.session_state.sa_selesai = saldo_awal_selesai()
    if not st.session_state.sa_selesai:
        sidebar_akun()
        halaman_saldo_awal()
        return

    muat_coa()      # akun tambahan (menu Tambah COA) dimuat sekali per sesi

    # 3) Aplikasi utama: menu sidebar dikelompokkan per bagian (st.navigation)
    halaman = st.navigation({
        "Ringkasan": [
            st.Page(halaman_dashboard, title="Dashboard", icon=":material/dashboard:",
                    url_path="dashboard", default=True),
        ],
        "Transaksi": [
            st.Page(halaman_pembelian, title="Pembelian", icon=":material/shopping_cart:",
                    url_path="pembelian"),
            st.Page(halaman_penjualan, title="Penjualan", icon=":material/sell:",
                    url_path="penjualan"),
            st.Page(halaman_daftar_harga, title="Daftar Harga", icon=":material/price_change:",
                    url_path="daftar-harga"),
            st.Page(halaman_penyesuaian, title="Penyesuaian", icon=":material/tune:",
                    url_path="penyesuaian"),
            st.Page(halaman_tarik_uang, title="Tarik Uang", icon=":material/payments:",
                    url_path="tarik-uang"),
        ],
        "Laporan": [
            st.Page(halaman_akses_cepat, title="Akses Cepat", icon=":material/receipt_long:",
                    url_path="akses-cepat"),
            st.Page(halaman_coa, title="Tambah COA", icon=":material/library_add:",
                    url_path="tambah-coa"),
            st.Page(halaman_laba_rugi, title="Laporan Laba Rugi", icon=":material/trending_up:",
                    url_path="laba-rugi"),
            st.Page(halaman_neraca, title="Laporan Neraca", icon=":material/balance:",
                    url_path="neraca"),
        ],
    })
    sidebar_akun()          # kartu akun muncul di bawah menu
    halaman.run()


# Streamlit menjalankan file ini sebagai "__main__", jadi main() otomatis dipanggil
if __name__ == "__main__":
    main()