"""
Parser PDF "Laporan Transaksi Finansial" BRI (giro / IBBIZ).

Kolom: Tanggal Transaksi (dd/mm/yy hh:mm:ss) | Uraian | Teller | Debet | Kredit | Saldo
Debet = uang keluar, Kredit = uang masuk. Ringkasan di halaman terakhir:
Saldo Awal | Total Debet | Total Kredit | Saldo Akhir.
"""

import re
from datetime import datetime

import pdfplumber

from pdf_ke_rekor_bca import kelompok_baris, teks_baris
from .common import (RE_ANGKA, angka, buat_transaksi, kelompokkan_ke_jangkar,
                     kolom_terdekat, nama_setelah_kepada, pilih_arah, pusat_kolom,
                     rapikan_nama, set_periode)

NAMA = "BRI"
RE_TGL = re.compile(r"^\d{2}/\d{2}/\d{2}$")
RE_JAM = re.compile(r"^\d{2}:\d{2}:\d{2}$")


def cocok(teks_halaman_1):
    return "LAPORAN TRANSAKSI FINANSIAL" in teks_halaman_1.upper()


def _header(pdf):
    page = pdf.pages[0]
    teks = page.extract_text() or ""
    words = page.extract_words(x_tolerance=2)
    info = {"Cabang": None, "Nama": rapikan_nama(nama_setelah_kepada(words)),
            "No Rekening": None, "Periode": None, "Mata Uang": None}

    m = re.search(r"No\.?\s*Rekening.*?:\s*(\d{6,})", teks, re.S)
    info["No Rekening"] = m.group(1) if m else None
    m = re.search(r"Valuta.*?:\s*([A-Z]{3})", teks, re.S)
    info["Mata Uang"] = m.group(1) if m else None
    m = re.search(r"Unit Kerja.*?:\s*(KC|KCP|Unit)\s+([^\n]+)", teks)
    info["Cabang"] = f"{m.group(1)} {m.group(2).strip()}" if m else None

    m = re.search(r"(\d{2}/\d{2}/\d{2})\s*-\s*(\d{2}/\d{2}/\d{2})", teks)
    if not m:
        raise ValueError("Periode transaksi tidak ditemukan di PDF BRI.")
    f = lambda s: datetime.strptime(s, "%d/%m/%y")      # noqa: E731
    set_periode(info, f(m.group(1)), f(m.group(2)))
    return info


def _nama_lawan(uraian):
    # "BY PPJ AN CV PRATAMA JAYA"  ->  CV PRATAMA JAYA
    m = re.search(r"\bAN\s+(.+)$", uraian)
    if m:
        return rapikan_nama(m.group(1))
    # "PRATAMA JAYA `CV-BANK DANAMON-BRI ..."  ->  CV PRATAMA JAYA
    m = re.match(r"^(.*?)\s*`(CV|PT|UD|PD)\b", uraian, re.I)
    if m and m.group(1).strip():
        return f"{m.group(2).upper()} {m.group(1).strip()}"
    return None


def _halaman(page, hasil):
    words = page.extract_words(x_tolerance=2)
    judul = [w for w in words if w["text"] == "Uraian"]
    if not judul:
        return
    h = min(judul, key=lambda w: w["top"])
    zona = [w for w in words if h["top"] - 6 <= w["top"] <= h["top"] + 22]
    bawah_header = max(w["bottom"] for w in zona)
    pusat = pusat_kolom(zona, {"debit": "Debet", "kredit": "Kredit", "saldo": "Saldo"})
    user = [w for w in zona if w["text"] == "User"]
    x_teller = (user[0]["x0"] if user else 300) - 3
    if len(pusat) < 3:
        raise ValueError("Header kolom BRI (Debet/Kredit/Saldo) tidak ditemukan.")

    body = [w for w in words if w["top"] > bawah_header]

    # potong di baris ringkasan "Saldo Awal Total Transaksi Debet ..."
    batas = None
    for b in kelompok_baris(body):
        if teks_baris(b).startswith("Saldo Awal Total"):
            batas = b[0]["top"] - 1
            # baris angka ringkasan = baris berikutnya yang berisi 4 angka
            for b2 in kelompok_baris([w for w in body if w["top"] > b[0]["top"] + 12]):
                angka_b2 = [w for w in b2 if RE_ANGKA.match(w["text"])]
                if len(angka_b2) >= 4:
                    sa, td, tk, sk = (angka(w["text"]) for w in angka_b2[:4])
                    hasil["ringkasan_umum"].update(
                        saldo_awal=sa, total_debit=td, total_kredit=tk, saldo_akhir=sk)
                    hasil["saldo_awal"] = sa
                    break
            break
    if batas is not None:
        body = [w for w in body if w["top"] < batas]

    jangkar = sorted((w for w in body if w["x0"] < 75 and RE_TGL.match(w["text"])),
                     key=lambda w: w["top"])
    for tgl_word, ws in zip(jangkar, kelompokkan_ke_jangkar(body, jangkar)):
        deskripsi, teller, nilai = [], [], {}
        for w in sorted(ws, key=lambda w: (round(w["top"]), w["x0"])):
            t = w["text"]
            if w is tgl_word or (w["x0"] < 110 and (RE_TGL.match(t) or RE_JAM.match(t))):
                continue
            if RE_ANGKA.match(t) and w["x0"] > x_teller:
                nilai[kolom_terdekat(w, pusat)] = angka(t)
            elif w["x0"] < x_teller:
                deskripsi.append(w)
            elif w["x0"] < pusat["debit"] - 25:
                teller.append(w)

        uraian = " ".join(teks_baris(b) for b in kelompok_baris(deskripsi)).strip()
        keluar, masuk = pilih_arah(nilai.get("debit"), nilai.get("kredit"))
        tgl = datetime.strptime(tgl_word["text"], "%d/%m/%y")
        id_teller = " ".join(w["text"] for w in sorted(teller, key=lambda w: w["x0"])) or None
        hasil["transaksi"].append(buat_transaksi(
            tgl, uraian, f"Teller {id_teller}" if id_teller else None,
            _nama_lawan(uraian), keluar, masuk, nilai.get("saldo"), detail=[uraian]))


def proses(file_obj):
    hasil = {"transaksi": [], "saldo_awal": None, "ringkasan": {}, "ringkasan_umum": {}}
    with pdfplumber.open(file_obj) as pdf:
        info = _header(pdf)
        for page in pdf.pages:
            _halaman(page, hasil)
    return info, hasil
