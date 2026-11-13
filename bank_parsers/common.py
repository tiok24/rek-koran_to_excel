"""
Fungsi bersama untuk parser bank (CIMB, Danamon, BRI) dan CSV.

Konvensi (sama dengan parser BCA):
    t["Debit"]  = uang KELUAR  (sisi debet di laporan bank)
    t["Kredit"] = uang MASUK   (sisi kredit di laporan bank)
Di Excel template REKOR keduanya DITUKAR (masuk -> kolom Debit, keluar -> kolom Kredit);
penukaran itu dilakukan oleh tulis_excel(), jadi parser tidak perlu memikirkannya.
"""

import re
from datetime import datetime

from pdf_ke_rekor_bca import BULAN, kelompok_baris, teks_baris  # noqa: F401  (dipakai ulang)

TOLERANSI = 0.01

# angka gaya Inggris: 1,234,567.89  (dipakai CIMB, Danamon, BRI)
RE_ANGKA = re.compile(r"^\d{1,3}(?:,\d{3})*\.\d{2}$")

NAMA_BULAN_ID = {v: k for k, v in BULAN.items()}          # 4 -> "APRIL"

# singkatan / nama bulan Inggris + Indonesia -> nomor
BULAN_SINGKAT = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "MEI": 5, "JUN": 6, "JUL": 7,
    "AUG": 8, "AGU": 8, "AGS": 8, "SEP": 9, "SEPT": 9, "OCT": 10, "OKT": 10, "NOV": 11,
    "DEC": 12, "DES": 12,
}
BULAN_PANJANG_EN = {
    "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6, "JULY": 7,
    "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12,
}


def nomor_bulan(teks):
    """'Feb' / 'FEBRUARI' / 'February' -> 2 (None bila tak dikenal)."""
    t = teks.strip().upper().rstrip(".")
    return BULAN_SINGKAT.get(t) or BULAN_PANJANG_EN.get(t) or BULAN.get(t) or BULAN_SINGKAT.get(t[:3])


def angka(teks):
    return float(teks.replace(",", ""))


def rapikan_nama(nama):
    """'ALAMINDO RAYA` PT' -> 'PT ALAMINDO RAYA' (bentuk badan usaha di belakang tanda `)."""
    if not nama:
        return nama
    nama = re.sub(r"\s+", " ", nama).strip()
    m = re.match(r"^(.*?)\s*`\s*(CV|PT|UD|PD|YAYASAN|KOPERASI|FA|FIRMA)\b\.?\s*$", nama, re.I)
    if m:
        return f"{m.group(2).upper()} {m.group(1).strip()}"
    return nama


def periode_teks(dt):
    """datetime -> 'APRIL 2026' (format yang dipahami bulan_tahun() milik modul BCA)."""
    return f"{NAMA_BULAN_ID[dt.month]} {dt.year}"


def set_periode(info, awal, akhir):
    info["periode_awal"], info["periode_akhir"] = awal, akhir
    info["Periode"] = periode_teks(awal)


def tahun_untuk_bulan(bulan, awal, akhir):
    """Tahun yang tepat untuk tanggal tanpa tahun (dd/mm) di dalam periode awal..akhir."""
    if awal.year == akhir.year:
        return awal.year
    return awal.year if bulan >= awal.month else akhir.year


def nama_file_hasil(info):
    """REKOR_<no rekening>_<tahun>-<bulan>.xlsx ; periode >1 bulan -> ..._<tahun>-<bln>_sd_<bln>."""
    a, b = info["periode_awal"], info["periode_akhir"]
    no = re.sub(r"[^\w\-]", "", info["No Rekening"] or "tanpa-nomor")
    if (a.year, a.month) == (b.year, b.month):
        return f"REKOR_{no}_{a.year}-{a.month:02d}.xlsx"
    if a.year == b.year:
        return f"REKOR_{no}_{a.year}-{a.month:02d}_sd_{b.month:02d}.xlsx"
    return f"REKOR_{no}_{a.year}-{a.month:02d}_sd_{b.year}-{b.month:02d}.xlsx"


# --------------------------------------------------------------------------
# Bantuan baca layout PDF
# --------------------------------------------------------------------------
def nama_setelah_kepada(words):
    """Nama pemilik rekening = baris pertama di bawah kata 'Kepada' (kotak alamat kiri atas)."""
    k = [w for w in words if w["text"] == "Kepada" and w["x0"] < 80]
    if not k:
        return None
    k = min(k, key=lambda w: w["top"])
    kiri = [w for w in words if w["x0"] < 330 and k["bottom"] - 1 < w["top"] < k["bottom"] + 40]
    baris = kelompok_baris(kiri)
    return teks_baris(baris[0]) if baris else None


def pusat_kolom(zona_header, nama):
    """{'debit': x_tengah, ...} dari kata header; `nama` = {'debit': 'Debet', ...}."""
    out = {}
    for kunci, teks in nama.items():
        c = [w for w in zona_header if w["text"] == teks]
        if c:
            out[kunci] = (c[0]["x0"] + c[0]["x1"]) / 2
    return out


def kolom_terdekat(word, pusat):
    """Kolom angka (debit/kredit/saldo) = pusat header terdekat dengan titik tengah angka.
    Berlaku untuk angka rata-kanan (BRI, Danamon) maupun rata-tengah (CIMB)."""
    x = (word["x0"] + word["x1"]) / 2
    return min(pusat, key=lambda k: abs(pusat[k] - x))


def kelompokkan_ke_jangkar(words, jangkar):
    """Tiap kata dimasukkan ke baris transaksi yang titik tengah vertikalnya paling dekat.
    Cocok untuk tabel yang isinya rata-tengah vertikal (keterangan multi-baris)."""
    mids = [(a["top"] + a["bottom"]) / 2 for a in jangkar]
    out = [[] for _ in jangkar]
    if not mids:
        return out
    for w in words:
        m = (w["top"] + w["bottom"]) / 2
        i = min(range(len(mids)), key=lambda k: abs(mids[k] - m))
        out[i].append(w)
    return out


def baris_dengan_kata(words, teks_awal, top_min=0):
    """Baris pertama (di bawah top_min) yang teksnya diawali `teks_awal`."""
    for b in kelompok_baris([w for w in words if w["top"] > top_min]):
        if teks_baris(b).startswith(teks_awal):
            return b
    return None


def buat_transaksi(tanggal, ket_bank, ket_lengkap, nama_lawan, keluar, masuk, saldo,
                   detail=None):
    """Dict transaksi dalam skema yang dipakai tulis_excel()/UI."""
    return {
        "Tanggal": tanggal,
        "Jenis Transaksi": ket_bank,
        "jenis_extra": [],
        "detail": detail or [],
        "Keterangan": " | ".join(detail or []),
        "CBG": None,
        "Debit": keluar,      # uang keluar
        "Kredit": masuk,      # uang masuk
        "Saldo": saldo,
        "ket_bank": ket_bank,
        "ket_lengkap": ket_lengkap,
        "nama_lawan": nama_lawan,
    }


def pilih_arah(nilai_debit, nilai_kredit):
    """Dari dua angka kolom debit/kredit -> (keluar, masuk); angka 0.00 dianggap kosong."""
    keluar = nilai_debit if nilai_debit else None
    masuk = nilai_kredit if nilai_kredit else None
    return keluar, masuk


# --------------------------------------------------------------------------
# Validasi
# --------------------------------------------------------------------------
def validasi_umum(hasil):
    """Periksa hasil terhadap ringkasan di dokumen (bila ada) dan terhadap saldo berjalan.

    hasil["ringkasan_umum"] boleh berisi: saldo_awal, total_debit (keluar),
    total_kredit (masuk), saldo_akhir. Kembalian: dict(ok, pesan[list], saldo_akhir, ...).
    """
    trx = hasil["transaksi"]
    rg = hasil.get("ringkasan_umum", {})
    pesan = []

    masuk = sum(t["Kredit"] or 0 for t in trx)
    keluar = sum(t["Debit"] or 0 for t in trx)

    saldo_awal = hasil.get("saldo_awal")
    if rg.get("saldo_awal") is not None and saldo_awal is not None \
            and abs(rg["saldo_awal"] - saldo_awal) > TOLERANSI:
        pesan.append("Saldo awal tidak sama dengan ringkasan di dokumen.")

    if rg.get("total_debit") is not None and abs(rg["total_debit"] - keluar) > TOLERANSI:
        pesan.append(f"Total uang keluar terbaca {keluar:,.2f}, di dokumen {rg['total_debit']:,.2f}.")
    if rg.get("total_kredit") is not None and abs(rg["total_kredit"] - masuk) > TOLERANSI:
        pesan.append(f"Total uang masuk terbaca {masuk:,.2f}, di dokumen {rg['total_kredit']:,.2f}.")

    n_keluar = sum(1 for t in trx if t["Debit"] is not None)
    n_masuk = sum(1 for t in trx if t["Kredit"] is not None)
    if rg.get("n_debit") is not None and rg["n_debit"] != n_keluar:
        pesan.append(f"Jumlah transaksi keluar terbaca {n_keluar}, di dokumen {rg['n_debit']}.")
    if rg.get("n_kredit") is not None and rg["n_kredit"] != n_masuk:
        pesan.append(f"Jumlah transaksi masuk terbaca {n_masuk}, di dokumen {rg['n_kredit']}.")

    tak_cocok = 0
    if saldo_awal is not None:
        berjalan = saldo_awal
        for t in trx:
            berjalan += (t["Kredit"] or 0) - (t["Debit"] or 0)
            if t["Saldo"] is not None and abs(berjalan - t["Saldo"]) > TOLERANSI:
                tak_cocok += 1
                berjalan = t["Saldo"]          # sinkronkan agar tidak berantai
        if tak_cocok:
            pesan.append(f"{tak_cocok} baris saldo berjalan tidak cocok.")

    saldo_terakhir = next((t["Saldo"] for t in reversed(trx) if t["Saldo"] is not None), None)
    if rg.get("saldo_akhir") is not None and saldo_terakhir is not None \
            and abs(rg["saldo_akhir"] - saldo_terakhir) > TOLERANSI:
        pesan.append("Saldo akhir tidak sama dengan ringkasan di dokumen.")

    # tidak ada satu pun pembanding -> tidak bisa dinyatakan cocok
    ada_pembanding = (any(v is not None for v in rg.values())
                      or any(t["Saldo"] is not None for t in trx))
    if not ada_pembanding:
        pesan.append("Tidak ada saldo/ringkasan pembanding di dokumen; hasil belum bisa diverifikasi.")

    return {
        "ok": not pesan,
        "bisa_diverifikasi": ada_pembanding,
        "pesan": pesan,
        "total_masuk": masuk,
        "total_keluar": keluar,
        "saldo_awal": saldo_awal,
        "saldo_akhir": rg.get("saldo_akhir", saldo_terakhir),
        "baris_saldo_tak_cocok": tak_cocok,
    }
