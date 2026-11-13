"""
Parser PDF "Transaction Inquiry Report" Bank Danamon (halaman landscape).

Kolom: Posting Date | Value Date | Transaction Branch | Reference Number | Description |
       Debit | Credit | Balance.     Debit = uang keluar, Credit = uang masuk.
Opening Balance ada di header halaman 1; baris "Total" (debit, credit, saldo akhir) di akhir.
Logo Danamon berupa gambar, jadi PDF dikenali dari judul "Transaction Inquiry Report".
"""

import re
from datetime import datetime

import pdfplumber

from pdf_ke_rekor_bca import kelompok_baris, teks_baris
from .common import (RE_ANGKA, angka, buat_transaksi, kelompokkan_ke_jangkar,
                     kolom_terdekat, nomor_bulan, pilih_arah, pusat_kolom, rapikan_nama,
                     set_periode)

NAMA = "DANAMON"
RE_PERIODE = re.compile(
    r"([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})\s*-\s*([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})")


def cocok(teks_halaman_1):
    t = teks_halaman_1.upper()
    return "TRANSACTION INQUIRY REPORT" in t and "POSTING DATE" in t


def _garis(words, label, x_maks=60):
    """Kata-kata di baris yang sama dengan `label` (kata di sisi kiri), selain label itu."""
    lbl = [w for w in words if w["text"] == label and w["x0"] < x_maks]
    if not lbl:
        return []
    top = lbl[0]["top"]
    return sorted((w for w in words if abs(w["top"] - top) < 3 and w is not lbl[0]),
                  key=lambda w: w["x0"])


def _header(pdf):
    page = pdf.pages[0]
    teks = page.extract_text() or ""
    words = page.extract_words(x_tolerance=2)
    info = {"Cabang": None, "Nama": None, "No Rekening": None, "Periode": None,
            "Mata Uang": None}

    m = re.search(r"Account Number\s*:\s*(\d+)", teks)
    info["No Rekening"] = m.group(1) if m else None

    # "Name : ALAMINDO RAYA` PT" (kata di kolom kiri, sebelum kolom "Print Date")
    nama = [w["text"] for w in _garis(words, "Name") if 100 < w["x0"] < 600 and w["text"] != ":"]
    info["Nama"] = rapikan_nama(" ".join(nama)) if nama else None
    mu = [w["text"] for w in _garis(words, "Currency") if 100 < w["x0"] < 600 and w["text"] != ":"]
    info["Mata Uang"] = mu[0] if mu else None

    m = RE_PERIODE.search(teks)
    if not m:
        raise ValueError("Periode tidak ditemukan di PDF Danamon.")
    b1, d1, y1, b2, d2, y2 = m.groups()
    set_periode(info,
                datetime(int(y1), nomor_bulan(b1), int(d1)),
                datetime(int(y2), nomor_bulan(b2), int(d2)))

    sa = [w for w in words if w["text"] == "Opening"]
    saldo_awal = None
    if sa:
        for w in words:
            if abs(w["top"] - sa[0]["top"]) < 3 and RE_ANGKA.match(w["text"]):
                saldo_awal = angka(w["text"])
    return info, saldo_awal


def _halaman(page, hasil):
    words = page.extract_words(x_tolerance=2)
    judul = [w for w in words if w["text"] == "Posting"]
    if not judul:
        return
    h = min(judul, key=lambda w: w["top"])
    zona = [w for w in words if h["top"] - 8 <= w["top"] <= h["top"] + 16]
    bawah_header = max(w["bottom"] for w in zona)
    pusat = pusat_kolom(zona, {"debit": "Debit", "kredit": "Credit", "saldo": "Balance"})
    if len(pusat) < 3:
        raise ValueError("Header kolom Danamon (Debit/Credit/Balance) tidak ditemukan.")
    desk = [w for w in zona if w["text"] == "Description"]
    x_desk = (desk[0]["x0"] if desk else 378) - 3
    hdr_debit = [w for w in zona if w["text"] == "Debit"][0]
    x_angka = hdr_debit["x0"] - 80

    body = [w for w in words if w["top"] > bawah_header]

    # baris Total (dan setelahnya) bukan transaksi
    for w in body:
        if w["text"] == "Total" and w["x0"] > 300:
            for a in body:
                if abs(a["top"] - w["top"]) < 3 and RE_ANGKA.match(a["text"]):
                    k = kolom_terdekat(a, pusat)
                    kunci = {"debit": "total_debit", "kredit": "total_kredit",
                             "saldo": "saldo_akhir"}[k]
                    hasil["ringkasan_umum"][kunci] = angka(a["text"])
            body = [b for b in body if b["top"] < w["top"] - 2]
            break

    # jangkar baris: tanggal posting = "15" "Jan" "2026" di kolom paling kiri
    jangkar, tanggal = [], {}
    for w in sorted(body, key=lambda w: w["top"]):
        if w["x0"] < 30 and re.fullmatch(r"\d{1,2}", w["text"]):
            bln = [b for b in body if abs(b["top"] - w["top"]) < 2 and 30 <= b["x0"] < 50
                   and nomor_bulan(b["text"])]
            thn = [b for b in body if abs(b["top"] - w["top"]) < 2 and 45 <= b["x0"] < 75
                   and re.fullmatch(r"\d{4}", b["text"])]
            if bln and thn:
                jangkar.append(w)
                tanggal[id(w)] = datetime(int(thn[0]["text"]), nomor_bulan(bln[0]["text"]),
                                          int(w["text"]))

    ref_x = [w for w in zona if w["text"] == "Reference"]
    x_ref = (ref_x[0]["x0"] if ref_x else 262) - 3

    for w_tgl, ws in zip(jangkar, kelompokkan_ke_jangkar(body, jangkar)):
        nilai, uraian, ref = {}, [], []
        for w in ws:
            t = w["text"]
            if RE_ANGKA.match(t) and w["x0"] >= x_angka:
                nilai[kolom_terdekat(w, pusat)] = angka(t)
            elif w["x0"] >= x_desk and w["x1"] <= hdr_debit["x0"] + 5:
                uraian.append(w)
            elif x_ref <= w["x0"] < x_desk:
                ref.append(w)
        baris = [teks_baris(b) for b in kelompok_baris(uraian)]
        deskripsi = " ".join(baris).strip()
        no_ref = " ".join(w["text"] for w in sorted(ref, key=lambda w: w["x0"]))
        if re.fullmatch(r"0+", no_ref or "0"):
            no_ref = None
        keluar, masuk = pilih_arah(nilai.get("debit"), nilai.get("kredit"))
        hasil["transaksi"].append(buat_transaksi(
            tanggal[id(w_tgl)], deskripsi, f"Ref {no_ref}" if no_ref else None, None,
            keluar, masuk, nilai.get("saldo"), detail=baris))


def proses(file_obj):
    hasil = {"transaksi": [], "saldo_awal": None, "ringkasan": {}, "ringkasan_umum": {}}
    with pdfplumber.open(file_obj) as pdf:
        info, saldo_awal = _header(pdf)
        hasil["saldo_awal"] = saldo_awal
        hasil["ringkasan_umum"]["saldo_awal"] = saldo_awal
        for page in pdf.pages:
            _halaman(page, hasil)
    return info, hasil
