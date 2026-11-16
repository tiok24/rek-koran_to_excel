"""
Parser PDF "Account Statement" Bank Mandiri (Kopra by Mandiri / rekening giro bisnis).

Kolom : Posting Date (dd Mmm yyyy, hh:mm:ss) | Remark | Reference No. | Debit | Credit | Balance
        Debit = uang KELUAR, Credit = uang MASUK (angka gaya Inggris 1,234,567.89).
Ringkasan di halaman 1: Opening Balance, No. of Debit, Total Amount Debited,
                        Closing Balance, No. of Credit, Total Amount Credited.

Ciri layout yang ditangani:
  * Satu transaksi bisa terdiri dari 1-5+ baris teks; angka (Debit/Credit/Balance) ada di
    baris tengah, jadi kolom Balance dipakai sebagai JANGKAR baris dan semua kata
    dikelompokkan ke jangkar terdekat secara vertikal.
  * Remark bisa terpotong ke halaman berikutnya (lanjutannya muncul di atas header tabel
    halaman berikut) -> ditempel ke transaksi terakhir halaman sebelumnya.
  * ID panjang (>= 21 karakter) dipotong di tengah token oleh PDF -> disambung tanpa spasi.
"""

import re
from datetime import datetime

import pdfplumber

from pdf_ke_rekor_bca import kelompok_baris, teks_baris
from .common import RE_ANGKA, angka, buat_transaksi, nomor_bulan, pilih_arah, set_periode

NAMA = "MANDIRI"

RE_TGL = re.compile(r"(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{4})")
RE_JAM = re.compile(r"\d{2}:\d{2}:\d{2}")
RE_BIC_NAMA = re.compile(r"\b[A-Z]{4}ID[A-Z]{2}/(.+)$")           # BNINIDJA/NAMA ...
RE_ID_BIFAST = re.compile(r"\d{8}[A-Z]{4}ID[A-Z]{2}\d{3}[A-Z]\d+")  # 20260114BMRIIDJA010O9...
PANJANG_TOKEN_TERPOTONG = 21


def cocok(teks_halaman_1):
    t = teks_halaman_1.upper()
    return "KOPRA BY MANDIRI" in t or (
        "ACCOUNT STATEMENT SUMMARY" in t and "POSTING DATE" in t and "REMARK" in t)


# --------------------------------------------------------------------------
# Header halaman 1
# --------------------------------------------------------------------------
def _tanggal(teks):
    m = RE_TGL.search(teks)
    bln = nomor_bulan(m.group(2)) if m else None
    if not bln:
        raise ValueError(f"Format tanggal tidak dikenali: '{teks}'.")
    return datetime(int(m.group(3)), bln, int(m.group(1)))


def _baris_di_bawah(baris, awalan):
    """(baris_label, baris_nilai) untuk baris yang diawali `awalan`."""
    for i, b in enumerate(baris[:-1]):
        if teks_baris(b).startswith(awalan):
            return b, baris[i + 1]
    return None, None


def _kolom(kata, batas):
    """Pisahkan kata ke kolom menurut daftar batas x0 (naik)."""
    out = [[] for _ in range(len(batas) + 1)]
    for w in kata:
        out[sum(1 for x in batas if w["x0"] >= x)].append(w)
    return [" ".join(w["text"] for w in sorted(k, key=lambda w: w["x0"])) for k in out]


def _x_label(label_baris, teks, setelah=0):
    c = [w for w in label_baris if w["x0"] > setelah and w["text"] == teks]
    return c[0]["x0"] if c else None


def _header(pdf):
    page = pdf.pages[0]
    baris = kelompok_baris(page.extract_words(x_tolerance=2))
    info = {"Cabang": None, "Nama": None, "No Rekening": None,
            "Periode": None, "Mata Uang": None}
    ringkas = {}

    lab, val = _baris_di_bawah(baris, "Account No.")
    if lab is None:
        raise ValueError("Header 'Account No.' tidak ditemukan.")
    x_nama = _x_label(lab, "Account", setelah=100)
    x_alias = _x_label(lab, "Alias")
    no, nama, _alias = _kolom(val, [x_nama - 5, x_alias - 5])
    info["No Rekening"], info["Nama"] = no.strip() or None, nama.strip() or None

    lab, val = _baris_di_bawah(baris, "Period")
    x_cur, x_cab = _x_label(lab, "Currency"), _x_label(lab, "Branch")
    periode, cur, cab = _kolom(val, [x_cur - 5, x_cab - 5])
    m = re.match(r"(.+?\d{4})\s*-\s*(.+?\d{4})", periode)
    if not m:
        raise ValueError("Periode tidak ditemukan di PDF Mandiri.")
    set_periode(info, _tanggal(m.group(1)), _tanggal(m.group(2)))
    info["Mata Uang"], info["Cabang"] = cur.strip() or None, cab.strip() or None

    for awalan, kunci in (("Opening Balance", ("saldo_awal", "n_debit", "total_debit")),
                          ("Closing Balance", ("saldo_akhir", "n_kredit", "total_kredit"))):
        _, v = _baris_di_bawah(baris, awalan)
        teks = [w["text"] for w in sorted(v, key=lambda w: w["x0"])] if v else []
        if len(teks) >= 3:
            ringkas[kunci[0]] = angka(teks[0])
            ringkas[kunci[1]] = int(teks[1])
            ringkas[kunci[2]] = angka(teks[2])
    return info, ringkas


# --------------------------------------------------------------------------
# Teks remark -> kolom Excel
# --------------------------------------------------------------------------
def _gabung(baris):
    """Sambung baris remark; token panjang yang terpotong di akhir BARIS MENTAH sebelumnya
    disambung tanpa spasi (yang dicek baris mentah, bukan hasil gabungan)."""
    out, sebelumnya = "", ""
    for b in baris:
        if not out:
            out = b
        elif len(sebelumnya.split()[-1]) >= PANJANG_TOKEN_TERPOTONG:
            out += b
        else:
            out += " " + b
        sebelumnya = b
    return out.strip()


def _huruf_besar(teks):
    """Deretan token awal yang berhuruf kapital & tanpa angka (nama pihak lawan)."""
    out = []
    for tok in teks.split():
        if tok != tok.upper() or any(c.isdigit() for c in tok):
            break
        out.append(tok)
    return " ".join(out) or None


def _tanpa_kode(teks):
    return re.sub(r"\s*\d{5}$", "", teks.strip())      # kode unit/transaksi: 14032, 99101, ...


def nama_lawan(t):
    m = re.search(r"MCM InhouseTrf\s+(?:DARI|KE)\s+(.+)$", t)
    if m:
        return _huruf_besar(_tanpa_kode(m.group(1)))
    m = re.search(r"MCM Outw CN\s+(.+?)\s+Clearing Fee", t)
    if m:
        return _huruf_besar(m.group(1))
    m = RE_BIC_NAMA.search(t)
    if m:
        return _huruf_besar(m.group(1))
    m = re.search(r"Transfer DARI\s+(.+)$", t)
    if m:
        return _huruf_besar(_tanpa_kode(m.group(1)))
    m = re.match(r"^(.+?)\s+-\s+\d{3}\b", t)                # "ERWIN ALBERTHO - 011 14032"
    if m:
        return m.group(1).strip()
    return None


def ket_bank(t):
    if "MCM Outw CN" in t:
        return "MCM Outw CN"
    if t.startswith("Clearing Fee"):
        return "Clearing Fee"
    m = re.search(r"MCM InhouseTrf\s+(DARI|KE)\b", t)
    if m:
        return f"MCM InhouseTrf {m.group(1)}"
    if "PRMA CR Transf" in t:
        return "PRMA CR Transf"
    if "Setor Tunai" in t:
        return "Setor Tunai"
    if "PEMBAYARAN TAGIHAN" in t:
        return "PEMBAYARAN TAGIHAN"
    if "Transfer DARI" in t:
        return "Transfer DARI"
    if RE_ID_BIFAST.search(t):                               # 2.500 per transfer keluar
        return "BI-FAST" if RE_BIC_NAMA.search(t) else "Biaya BI-FAST"
    m = re.match(r"^(Biaya Adm|Bunga|Pajak)\b", t)
    if m:
        return m.group(1)
    if re.match(r"^.+?\s+-\s+\d{3}\b", t):
        return "Transfer"
    return t.split(" ")[0] if t else ""


def _ket_lengkap(baris_remark, referensi):
    teks = _gabung([b for b in baris_remark if b.strip() not in ("", "-")])
    if referensi and referensi not in teks.replace(" ", ""):
        teks = f"{teks} | Ref: {referensi}" if teks else f"Ref: {referensi}"
    return teks


# --------------------------------------------------------------------------
# Halaman tabel
# --------------------------------------------------------------------------
def _kolom_angka(w, kanan):
    return min(kanan, key=lambda k: abs(kanan[k] - w["x1"]))


def _halaman(page, hasil, halaman_pertama):
    words = page.extract_words(x_tolerance=2)
    judul = [w for w in words if w["text"] == "Posting" and w["x0"] < 100]
    if not judul:
        return
    h = min(judul, key=lambda w: w["top"])
    zona = [w for w in words if abs(w["top"] - h["top"]) < 4]

    def hdr(teks):
        c = [w for w in zona if w["text"] == teks]
        return c[0] if c else None

    if not all(hdr(k) for k in ("Remark", "Reference", "Debit", "Credit", "Balance")):
        raise ValueError("Header tabel Mandiri (Remark/Reference/Debit/Credit/Balance) tidak lengkap.")
    x_remark, x_ref = hdr("Remark")["x0"] - 3, hdr("Reference")["x0"] - 3
    kanan = {"debit": hdr("Debit")["x1"], "kredit": hdr("Credit")["x1"],
             "saldo": hdr("Balance")["x1"]}
    bawah_header = max(w["bottom"] for w in zona)

    # lanjutan remark transaksi terakhir halaman sebelumnya (di atas header, di bawah banner)
    if not halaman_pertama and hasil["transaksi"]:
        lanjut = [w for w in words if 60 < w["top"] < h["top"] - 3 and x_remark <= w["x0"] < x_ref]
        tambahan = [teks_baris(b) for b in kelompok_baris(lanjut)]
        if tambahan:
            t = hasil["transaksi"][-1]
            t["detail"] = t["detail"] + tambahan
            t["_remark"] = _gabung(t["detail"])
            t["ket_lengkap"] = _ket_lengkap(t["detail"], t["_ref"])
            t["nama_lawan"] = nama_lawan(t["_remark"])
            t["ket_bank"] = t["Jenis Transaksi"] = ket_bank(t["_remark"])

    body = [w for w in words if w["top"] > bawah_header]
    for b in kelompok_baris(body):                           # buang footer halaman
        if teks_baris(b).startswith("For further questions"):
            body = [w for w in body if w["top"] < b[0]["top"] - 1]
            break

    jangkar = sorted((w for w in body if RE_ANGKA.match(w["text"])
                      and abs(w["x1"] - kanan["saldo"]) < 15), key=lambda w: w["top"])
    if not jangkar:
        return
    mids = [(a["top"] + a["bottom"]) / 2 for a in jangkar]
    grup = [[] for _ in jangkar]
    for w in body:
        m = (w["top"] + w["bottom"]) / 2
        grup[min(range(len(mids)), key=lambda i: abs(mids[i] - m))].append(w)

    for a, ws in zip(jangkar, grup):
        tgl, remark, ref, nilai = [], [], [], {}
        for w in ws:
            if w["x0"] < x_remark:
                tgl.append(w)
            elif RE_ANGKA.match(w["text"]) and w["x0"] >= x_ref:
                nilai[_kolom_angka(w, kanan)] = angka(w["text"])
            elif w["x0"] < x_ref:
                remark.append(w)
            else:
                ref.append(w)

        teks_tgl = " ".join(w["text"] for w in sorted(tgl, key=lambda w: (round(w["top"]), w["x0"])))
        tanggal = _tanggal(teks_tgl)
        baris_remark = [teks_baris(b) for b in kelompok_baris(remark)]
        baris_remark = [b for b in baris_remark if b.strip() not in ("", "-")]
        referensi = "".join(w["text"] for w in sorted(ref, key=lambda w: (round(w["top"]), w["x0"])))
        referensi = "" if referensi == "-" else referensi

        t_remark = _gabung(baris_remark)
        keluar, masuk = pilih_arah(nilai.get("debit"), nilai.get("kredit"))
        t = buat_transaksi(tanggal, ket_bank(t_remark), _ket_lengkap(baris_remark, referensi),
                           nama_lawan(t_remark), keluar, masuk, nilai.get("saldo"),
                           detail=baris_remark)
        t["_remark"], t["_ref"] = t_remark, referensi
        t["Waktu"] = RE_JAM.search(teks_tgl).group(0) if RE_JAM.search(teks_tgl) else None
        hasil["transaksi"].append(t)


def proses(file_obj):
    hasil = {"transaksi": [], "saldo_awal": None, "ringkasan": {}, "ringkasan_umum": {}}
    with pdfplumber.open(file_obj) as pdf:
        info, ringkas = _header(pdf)
        hasil["saldo_awal"] = ringkas.get("saldo_awal")
        hasil["ringkasan_umum"] = ringkas
        for i, page in enumerate(pdf.pages):
            _halaman(page, hasil, halaman_pertama=(i == 0))
    return info, hasil
