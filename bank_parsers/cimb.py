"""
Parser PDF "Laporan Transaksi / Account Statement" CIMB Niaga (giro perusahaan).

Kolom: Tgl. Txn | Tgl. Valuta | Uraian (multi-baris) | No. Cek/BG | Debet | Kredit | Saldo
Angka di CIMB rata-tengah terhadap header kolomnya. Baris SALDO AWAL di atas transaksi
pertama, lalu baris Total (debet, kredit) dan SALDO AKHIR di halaman terakhir.
Keterangan Bank = baris pertama uraian; Keterangan 1 = baris sisanya; Keterangan 2 = nama
lawan transaksi bila bisa dikenali (BI-FAST, SKN, overbooking).
"""

import re
from datetime import datetime

import pdfplumber

from pdf_ke_rekor_bca import kelompok_baris, teks_baris
from .common import (RE_ANGKA, angka, buat_transaksi,
                     kolom_terdekat, nama_setelah_kepada, nomor_bulan, pilih_arah,
                     pusat_kolom, rapikan_nama, set_periode, tahun_untuk_bulan)

NAMA = "CIMB"
RE_TGL = re.compile(r"^\d{2}/\d{2}$")
# jarak vertikal (pt) antar-baris di dalam satu transaksi ~8,4; antar-transaksi ~12,7
JEDA_ANTAR_TRANSAKSI = 10.6
RE_PERIODE = re.compile(r"(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})\s*-\s*(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})")


def cocok(teks_halaman_1):
    return "CIMB NIAGA" in teks_halaman_1.upper()


def _header(pdf):
    page = pdf.pages[0]
    teks = page.extract_text() or ""
    words = page.extract_words(x_tolerance=2)
    info = {"Cabang": None, "Nama": rapikan_nama(nama_setelah_kepada(words)),
            "No Rekening": None, "Periode": None, "Mata Uang": None}

    m = re.search(r"Account Number\s*:\s*(\d+)", teks)
    info["No Rekening"] = m.group(1) if m else None
    m = re.search(r"Currency\s*:\s*([A-Z]{3})", teks)
    info["Mata Uang"] = m.group(1) if m else None

    m = RE_PERIODE.search(teks[teks.find("Periode"):] if "Periode" in teks else teks)
    if not m:
        raise ValueError("Periode tidak ditemukan di PDF CIMB.")
    d1, b1, y1, d2, b2, y2 = m.groups()
    awal = datetime(int(y1), nomor_bulan(b1), int(d1))
    akhir = datetime(int(y2), nomor_bulan(b2), int(d2))
    set_periode(info, awal, akhir)
    return info


def _nama_lawan(detail):
    for baris in detail:
        # BI-FAST: "BFS FR BMRIIDJA UN ARIEF BRAMANTY"
        m = re.search(r"\bFR\s+\S+\s+UN\s+(.+)$", baris)
        if m:
            return rapikan_nama(m.group(1))
        # SKN/RTGS masuk: "BDINIDJA/AP. SURYA KENCANA`CV"
        m = re.match(r"^[A-Z]{4}ID[A-Z0-9]{2}(?:[A-Z0-9]{3})?/(.+)$", baris)
        if m:
            return rapikan_nama(m.group(1))
        # overbooking / transfer antar rekening: "BizChannel TRF TO BERDIKARI JAYA"
        m = re.search(r"\bTRF\s+(?:TO|FR)\s+(.+)$", baris, re.I)
        if m:
            return rapikan_nama(m.group(1))
    return None


def _tahun(tgl_txt, info):
    d, m = map(int, tgl_txt.split("/"))
    return datetime(tahun_untuk_bulan(m, info["periode_awal"], info["periode_akhir"]), m, d)


def _halaman(page, info, hasil):
    words = page.extract_words(x_tolerance=2)
    judul = [w for w in words if w["text"] == "Uraian"]
    if not judul:
        return
    h = min(judul, key=lambda w: w["top"])
    zona = [w for w in words if h["top"] - 6 <= w["top"] <= h["top"] + 18]
    bawah_header = max(w["bottom"] for w in zona)
    pusat = pusat_kolom(zona, {"debit": "Debet", "kredit": "Kredit", "saldo": "Saldo"})
    if len(pusat) < 3:
        raise ValueError("Header kolom CIMB (Debet/Kredit/Saldo) tidak ditemukan.")
    x_uraian = h["x0"] - 3
    cek_hdr = [w for w in zona if w["text"] in ("No.", "Cek/BG") and w["x0"] > 250]
    x_cek = (min(w["x0"] for w in cek_hdr) if cek_hdr else 300) - 3

    body = [w for w in words if w["top"] > bawah_header]

    # buang catatan kaki ("Terima kasih ...", nomor halaman)
    kaki = [w for w in body if w["text"] == "Terima" and w["x0"] < 40]
    if kaki:
        body = [w for w in body if w["top"] < min(k["top"] for k in kaki) - 2]

    # baris khusus: SALDO AWAL / Total / SALDO AKHIR (dikeluarkan dari baris transaksi)
    def angka_di_baris(top, tol=4):
        return [w for w in body if abs(w["top"] - top) < tol and RE_ANGKA.match(w["text"])]

    batas_bawah = None
    for b in kelompok_baris(body):
        t = teks_baris(b)
        top = b[0]["top"]
        if t.startswith("SALDO AWAL"):
            a = angka_di_baris(top)
            if a:
                hasil["saldo_awal"] = angka(a[-1]["text"])
                hasil["ringkasan_umum"]["saldo_awal"] = hasil["saldo_awal"]
            body = [w for w in body if abs(w["top"] - top) >= 4]
        elif t.startswith("Total") and b[0]["x0"] < 40:
            for w in angka_di_baris(top):
                k = kolom_terdekat(w, pusat)
                if k == "debit":
                    hasil["ringkasan_umum"]["total_debit"] = angka(w["text"])
                elif k == "kredit":
                    hasil["ringkasan_umum"]["total_kredit"] = angka(w["text"])
            batas_bawah = top - 2 if batas_bawah is None else min(batas_bawah, top - 2)
        elif t.startswith("SALDO AKHIR"):
            a = angka_di_baris(top)
            if a:
                hasil["ringkasan_umum"]["saldo_akhir"] = angka(a[-1]["text"])
            batas_bawah = top - 2 if batas_bawah is None else min(batas_bawah, top - 2)
    if batas_bawah is not None:
        body = [w for w in body if w["top"] < batas_bawah]

    jangkar = sorted((w for w in body if w["x0"] < 40 and RE_TGL.match(w["text"])),
                     key=lambda w: w["top"])
    if not jangkar:
        return
    mids = [(a["top"] + a["bottom"]) / 2 for a in jangkar]

    # angka (debet/kredit/saldo) sebaris dengan tanggal -> baris terdekat
    nilai = [{} for _ in jangkar]
    for w in body:
        if RE_ANGKA.match(w["text"]) and w["x0"] > x_cek + 20:
            m = (w["top"] + w["bottom"]) / 2
            i = min(range(len(mids)), key=lambda k: abs(mids[k] - m))
            nilai[i][kolom_terdekat(w, pusat)] = angka(w["text"])

    # uraian: kelompokkan menjadi BLOK memakai jeda antar-transaksi (lebih lebar daripada
    # jarak antar-baris di dalam satu transaksi), lalu pasang tiap blok ke tanggal yang
    # jatuh di rentang vertikalnya. Pemasangan "baris terdekat" saja keliru untuk blok
    # 5 baris yang bersebelahan dengan blok 2 baris.
    baris_uraian = kelompok_baris(
        [w for w in body if x_uraian <= w["x0"] < x_cek and not RE_ANGKA.match(w["text"])])
    blok, cur = [], []
    for b in baris_uraian:
        if cur and b[0]["top"] - cur[-1][0]["top"] > JEDA_ANTAR_TRANSAKSI:
            blok.append(cur)
            cur = []
        cur.append(b)
    if cur:
        blok.append(cur)

    uraian = [[] for _ in jangkar]
    for bl in blok:
        atas, bawah = bl[0][0]["top"] - 2, max(w["bottom"] for w in bl[-1]) + 2

        def jarak(m):
            return 0 if atas <= m <= bawah else min(abs(m - atas), abs(m - bawah))
        i = min(range(len(mids)), key=lambda k: jarak(mids[k]))
        uraian[i].extend(teks_baris(b) for b in bl)

    for w_tgl, baris, nl in zip(jangkar, uraian, nilai):
        baris = [b for b in baris if b.strip()]
        if not baris:
            continue
        keluar, masuk = pilih_arah(nl.get("debit"), nl.get("kredit"))
        detail = baris[1:]
        hasil["transaksi"].append(buat_transaksi(
            _tahun(w_tgl["text"], info), baris[0],
            " ".join(detail) or None, _nama_lawan(detail),
            keluar, masuk, nl.get("saldo"), detail=detail))


def proses(file_obj):
    hasil = {"transaksi": [], "saldo_awal": None, "ringkasan": {}, "ringkasan_umum": {}}
    with pdfplumber.open(file_obj) as pdf:
        info = _header(pdf)
        for page in pdf.pages:
            _halaman(page, info, hasil)
    return info, hasil
