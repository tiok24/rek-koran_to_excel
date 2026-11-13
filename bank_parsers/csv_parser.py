"""
CSV mutasi -> struktur yang sama dengan hasil parser PDF, sehingga bisa ditulis ke template REKOR.

CSV dari tiap bank/aplikasi berbeda, jadi modul ini tidak mengunci satu format. Alurnya:
    1. muat_csv()      : deteksi encoding, pemisah (, ; tab |) dan baris header.
    2. tebak_pemetaan(): tebak kolom tanggal/keterangan/debit/kredit/saldo dari nama header.
    3. konversi_csv()  : ubah baris menjadi transaksi (angka & tanggal diurai otomatis).
Pemetaan hasil tebakan bisa dikoreksi pengguna (di UI Streamlit) sebelum dikonversi.

Konvensi: kolom "debit" di CSV = uang KELUAR dan "kredit" = uang MASUK (sudut pandang bank).
Bila CSV Anda kebalikannya, pakai opsi tukar_arah=True.
"""

import csv
import io
import re
from datetime import datetime

from .common import buat_transaksi, nomor_bulan, set_periode

# kata kunci nama kolom (huruf kecil, tanpa tanda baca), urut dari yang paling spesifik
ALIAS = {
    "tanggal": ["tanggal transaksi", "tgl transaksi", "tanggal", "tgl", "posting date",
                "transaction date", "txn date", "trx date", "date", "waktu"],
    "keterangan": ["uraian transaksi", "keterangan", "uraian", "deskripsi", "description",
                   "transaction description", "remark", "remarks", "narasi", "berita",
                   "detail", "catatan", "note"],
    "debit": ["mutasi debet", "mutasi debit", "debet", "debit", "uang keluar", "keluar",
              "withdrawal", "db", "dr"],
    "kredit": ["mutasi kredit", "kredit", "credit", "uang masuk", "masuk", "deposit", "cr"],
    "saldo": ["saldo akhir", "running balance", "saldo", "balance"],
    "jumlah": ["jumlah", "amount", "nominal", "mutasi", "nilai"],
    "tipe": ["db cr", "cr db", "d k", "dk", "tipe transaksi", "tipe", "type", "jenis"],
}
KUNCI = list(ALIAS)


def _norm(teks):
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", str(teks).lower())).strip()


# --------------------------------------------------------------------------
# Membaca berkas
# --------------------------------------------------------------------------
def _decode(data):
    for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return data.decode("utf-8", errors="replace")


def _deteksi_pemisah(baris):
    """Pemisah dengan jumlah kolom paling konsisten (>1) pada 30 baris pertama."""
    terbaik, skor_terbaik = ",", -1
    for p in (",", ";", "\t", "|"):
        hitung = [len(next(csv.reader([b], delimiter=p))) for b in baris[:30] if b.strip()]
        if not hitung or max(hitung) < 2:
            continue
        modus = max(set(hitung), key=hitung.count)
        skor = hitung.count(modus) * (1 if modus > 1 else 0) + modus * 0.01
        if skor > skor_terbaik:
            terbaik, skor_terbaik = p, skor
    return terbaik


def _skor_header(sel):
    """Berapa jenis kolom yang dikenali pada satu baris."""
    n = {_norm(s) for s in sel}
    return sum(1 for k in KUNCI if any(a in n for a in ALIAS[k]))


def muat_csv(data, pemisah=None):
    """bytes -> (header[list], baris[list[list]], pemisah). Baris sebelum header (judul
    laporan bank, dll) dibuang; baris kosong dibuang."""
    teks = _decode(data)
    garis = [g for g in teks.splitlines()]
    pemisah = pemisah or _deteksi_pemisah(garis)
    semua = [r for r in csv.reader(garis, delimiter=pemisah) if any(c.strip() for c in r)]
    if not semua:
        raise ValueError("CSV kosong.")
    # baris header = baris (dalam 40 pertama) dengan kolom dikenali terbanyak
    idx = max(range(min(40, len(semua))), key=lambda i: (_skor_header(semua[i]), -i))
    if _skor_header(semua[idx]) < 2:
        idx = 0
    header = [h.strip() or f"kolom_{i + 1}" for i, h in enumerate(semua[idx])]
    lebar = len(header)
    baris = [(r + [""] * lebar)[:lebar] for r in semua[idx + 1:]]
    return header, baris, pemisah


def tebak_pemetaan(header):
    """{'tanggal': nama_kolom|None, ...}. Tiap kolom hanya dipakai sekali."""
    peta, dipakai = {}, set()
    normal = [_norm(h) for h in header]
    for kunci in KUNCI:
        peta[kunci] = None
        for alias in ALIAS[kunci]:
            ada = [i for i, n in enumerate(normal) if n == alias and i not in dipakai]
            if not ada:
                ada = [i for i, n in enumerate(normal)
                       if alias in n.split() and i not in dipakai]
            if ada:
                peta[kunci] = header[ada[0]]
                dipakai.add(ada[0])
                break
    # satu kolom "jumlah" hanya dipakai bila debit/kredit terpisah tidak ada
    if peta["debit"] and peta["kredit"]:
        peta["jumlah"] = peta["tipe"] = None
    return peta


# --------------------------------------------------------------------------
# Mengurai angka & tanggal
# --------------------------------------------------------------------------
def parse_angka(teks, format_angka="auto"):
    """'1.234.567,89' / '1,234,567.89' / '(1.000,00)' / '-500' / 'Rp 2.000' -> float | None.

    format_angka: 'auto', 'id' (1.234,56) atau 'en' (1,234.56)."""
    if teks is None:
        return None
    s = str(teks).strip()
    if s in ("", "-", "nan", "NaN", "None"):
        return None
    negatif = s.startswith("-") or s.endswith("-") or (s.startswith("(") and s.endswith(")"))
    s = re.sub(r"[^\d.,]", "", s)
    if not s or not re.search(r"\d", s):
        return None

    if format_angka == "id":
        s = s.replace(".", "").replace(",", ".")
    elif format_angka == "en":
        s = s.replace(",", "")
    else:
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")      # 1.234,56
            else:
                s = s.replace(",", "")                        # 1,234.56
        elif "," in s:
            bagian = s.split(",")
            if len(bagian) == 2 and len(bagian[1]) in (1, 2):
                s = bagian[0] + "." + bagian[1]               # 1234,56
            else:
                s = s.replace(",", "")                        # 1,234,567
        elif "." in s:
            bagian = s.split(".")
            if len(bagian) > 2 or len(bagian[1]) == 3:
                s = s.replace(".", "")                        # 1.234.567 / 1.234 (ribuan)
    try:
        nilai = float(s)
    except ValueError:
        return None
    return -nilai if negatif else nilai


def parse_tanggal(teks, hari_dulu=True):
    """Beberapa format umum -> datetime | None. Jam diabaikan."""
    if teks is None:
        return None
    s = str(teks).strip()
    if not s:
        return None
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)                 # 2026-04-09
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})", s)             # 09/04/2026
        if m:
            a, b, y = map(int, m.groups())
            d, mo = (a, b) if hari_dulu else (b, a)
            y += 2000 if y < 100 else 0
        else:
            m = re.match(r"^(\d{1,2})[\s\-/.,]*([A-Za-z]{3,9})[\s\-/.,]*(\d{2,4})", s)  # 09 Apr 2026
            if not m:
                return None
            d, nama, y = int(m.group(1)), m.group(2), int(m.group(3))
            mo = nomor_bulan(nama)
            if not mo:
                return None
            y += 2000 if y < 100 else 0
    try:
        return datetime(y, mo, d)
    except ValueError:
        return None


def _arah_dari_tipe(teks):
    t = str(teks).strip().upper()
    if t in ("DB", "D", "DR", "DEBIT", "DEBET") or t.startswith(("DEB", "DB")):
        return "keluar"
    if t in ("CR", "C", "K", "KR", "KREDIT", "CREDIT") or t.startswith(("CRE", "KRE", "CR")):
        return "masuk"
    return None


# --------------------------------------------------------------------------
# Konversi
# --------------------------------------------------------------------------
def konversi_csv(header, baris, pemetaan, *, nama="", no_rekening="", format_angka="auto",
                 tukar_arah=False, saldo_awal=None, hari_dulu=True):
    """Kembalikan (info, hasil) seperti parser PDF. Baris tanpa tanggal valid (judul,
    total, catatan kaki) dilewati dan dihitung di hasil['dilewati']."""
    if not pemetaan.get("tanggal"):
        raise ValueError("Kolom tanggal belum dipilih.")
    ada_dk = pemetaan.get("debit") or pemetaan.get("kredit")
    if not ada_dk and not pemetaan.get("jumlah"):
        raise ValueError("Pilih kolom debit/kredit, atau satu kolom jumlah.")

    ix = {h: i for i, h in enumerate(header)}
    ambil = lambda r, k: r[ix[pemetaan[k]]] if pemetaan.get(k) in ix else None     # noqa: E731

    transaksi, dilewati = [], 0
    for r in baris:
        tgl = parse_tanggal(ambil(r, "tanggal"), hari_dulu)
        if tgl is None:
            dilewati += 1
            continue
        keluar = masuk = None
        if ada_dk:
            keluar = parse_angka(ambil(r, "debit"), format_angka)
            masuk = parse_angka(ambil(r, "kredit"), format_angka)
            keluar = abs(keluar) if keluar else None
            masuk = abs(masuk) if masuk else None
        else:
            jml = parse_angka(ambil(r, "jumlah"), format_angka)
            if jml:
                arah = _arah_dari_tipe(ambil(r, "tipe")) if pemetaan.get("tipe") else None
                if arah is None:
                    arah = "keluar" if jml < 0 else "masuk"
                if arah == "keluar":
                    keluar = abs(jml)
                else:
                    masuk = abs(jml)
        if tukar_arah:
            keluar, masuk = masuk, keluar
        if not keluar and not masuk:
            dilewati += 1
            continue
        ket = re.sub(r"\s+", " ", str(ambil(r, "keterangan") or "")).strip()
        transaksi.append(buat_transaksi(
            tgl, ket, None, None, keluar, masuk,
            parse_angka(ambil(r, "saldo"), format_angka), detail=[ket] if ket else []))

    if not transaksi:
        raise ValueError("Tidak ada baris transaksi yang terbaca. Periksa pemetaan kolom, "
                         "format angka, dan format tanggal.")

    # CSV "terbaru di atas" -> balik supaya kronologis
    terbalik = len(transaksi) > 1 and transaksi[0]["Tanggal"] > transaksi[-1]["Tanggal"]
    if terbalik:
        transaksi.reverse()

    if saldo_awal is None:
        t0 = transaksi[0]
        saldo_awal = (t0["Saldo"] - (t0["Kredit"] or 0) + (t0["Debit"] or 0)
                      if t0["Saldo"] is not None else 0.0)

    info = {"Cabang": None, "Nama": nama or "", "No Rekening": no_rekening or "CSV",
            "Periode": None, "Mata Uang": None}
    set_periode(info, transaksi[0]["Tanggal"], transaksi[-1]["Tanggal"])
    hasil = {"transaksi": transaksi, "saldo_awal": saldo_awal, "ringkasan": {},
             "ringkasan_umum": {}, "dilewati": dilewati, "urutan_dibalik": terbalik}
    return info, hasil
