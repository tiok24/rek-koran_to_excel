"""
Parser PDF "Account Statement" BNI (rekening giro / BNI Direct).

Kolom : Posting Date | Effective Date | Branch | Journal | Transaction Description |
        Amount | DB/CR | Balance
        DB/CR: D = uang KELUAR (debet), K = uang MASUK (kredit). Angka gaya Inggris 1,234,567.89.
Ringkasan di halaman terakhir: Ending Balance, Total Debet (n, jumlah), Total Credit (n, jumlah).
Saldo awal = baris "Ledger Balance" di halaman 1.

Ciri layout yang ditangani:
  * Semua teks tercetak dua kali pada posisi yang sama (efek "tebal" palsu) -> dedupe_chars().
  * Pada beberapa baris ada satu karakter nyasar di dalam angka (mis. 4,862.00 terbaca 4,8602.00).
    Kandidat angka dibuat dari karakter yang tumpang tindih, lalu dipilih yang cocok dengan
    selisih saldo.
  * Tinggi baris bervariasi (keterangan 1-6 baris, rata tengah vertikal), jadi batas baris diambil
    dari kotak sel latar kolom "Posting Date", bukan dari jarak ke tanggal terdekat. Bila kotak
    tidak terbaca, parser berhenti dengan pesan error (bukan menebak).
  * Keterangan satu transaksi bisa terpotong ke halaman berikutnya (lanjutan muncul di atas
    tabel halaman berikut) -> ditempel ke transaksi terakhir halaman sebelumnya.
  * Baris terakhir laporan tidak mencetak saldo; diisi dari "Ending Balance" bila saldo
    berjalan cocok.
"""

import itertools
import re
from datetime import datetime

import pdfplumber

from pdf_ke_rekor_bca import kelompok_baris, teks_baris
from .common import (RE_ANGKA, TOLERANSI, angka, buat_transaksi, nomor_bulan, set_periode)

NAMA = "BNI"

RE_TGL = re.compile(r"^\d{2}/\d{2}/\d{4}$")
RE_PERIODE = re.compile(r"(\d{1,2})-([A-Za-z]{3,})-(\d{2,4})\s*-\s*(\d{1,2})-([A-Za-z]{3,})-(\d{2,4})")
RE_ANGKA_LONGGAR = re.compile(r"^[\d,\.]+$")
RE_REK_PINDAH = re.compile(r"^PEMINDAHAN\s+(?:DARI|KE)\s+(\d+)\s*(.*)$", re.I)
HONORIFIK = {"SDRI", "SDR", "SDRA", "BPK", "BAPAK", "IBU", "TN", "NY", "NN", "SAUDARA", "SAUDARI"}
PANJANG_NAMA_ECHANNEL = 24      # kolom nama di keterangan TRF/PAY/TOP-UP ECHANNEL lebarnya 24 karakter


def cocok(teks_halaman_1):
    """Logo BNI berupa gambar (tidak ada teks 'BNI'), jadi dikenali dari judul dan kolom tabel."""
    t = teks_halaman_1.upper()
    return ("ACCOUNT STATEMENT" in t and "ACCOUNT STATEMENT SUMMARY" not in t
            and "LEDGER BALANCE" in t and "JOURNAL" in t and "DB/CR" in t
            and "EFFECTIVE DATE" in t)


# --------------------------------------------------------------------------
# Angka
# --------------------------------------------------------------------------
def _kandidat_angka(page, w):
    """Semua angka sah yang mungkin dibentuk dari karakter di kotak kata `w`.
    Karakter yang tumpang tindih di posisi sama dan berbeda (mis. '2' dan '0') menjadi pilihan."""
    cs = [c for c in page.chars
          if c["x0"] >= w["x0"] - 0.6 and c["x1"] <= w["x1"] + 0.6 and abs(c["top"] - w["top"]) < 3]
    cs.sort(key=lambda c: c["x0"])
    klaster = []
    for c in cs:
        if klaster and abs(c["x0"] - klaster[-1][0]) < 0.8:
            if c["text"] not in klaster[-1][1]:
                klaster[-1][1].append(c["text"])
        else:
            klaster.append([c["x0"], [c["text"]]])
    if not klaster or len(klaster) > 20:
        return []
    out = []
    for kombinasi in itertools.islice(itertools.product(*[k[1] for k in klaster]), 256):
        s = "".join(kombinasi)
        if RE_ANGKA.match(s) and angka(s) not in out:
            out.append(angka(s))
    return out


def _nilai(page, w):
    """-> list kandidat angka untuk kata `w` (1 elemen bila tidak ambigu, kosong bila gagal)."""
    if RE_ANGKA.match(w["text"]):
        return [angka(w["text"])]
    return _kandidat_angka(page, w)


# --------------------------------------------------------------------------
# Header halaman 1
# --------------------------------------------------------------------------
def _tanggal_periode(hari, bulan, tahun):
    bln = nomor_bulan(bulan)
    if not bln:
        raise ValueError(f"Nama bulan periode tidak dikenali: '{bulan}'.")
    th = int(tahun)
    return datetime(th + 2000 if th < 100 else th, bln, int(hari))


def _header(page):
    teks = page.extract_text() or ""
    words = page.extract_words(x_tolerance=2)
    info = {"Cabang": None, "Nama": None, "No Rekening": None, "Periode": None, "Mata Uang": None}

    m = re.search(r"Account\s*No\.?\s*:\s*(\d{6,})", teks)
    info["No Rekening"] = m.group(1) if m else None
    m = re.search(r"\(([A-Z]{3})\)", teks)
    info["Mata Uang"] = m.group(1) if m else None

    # Nama = baris pertama kotak kiri, di bawah judul "ACCOUNT STATEMENT" dan kiri label "Account No."
    judul = [w for w in words if w["text"] == "STATEMENT"]
    label = [w for w in words if w["text"] == "Account"]
    if judul and label:
        x_batas = min(w["x0"] for w in label) - 5
        kiri = [w for w in words if w["x1"] < x_batas and w["top"] > judul[0]["bottom"] + 5]
        baris = kelompok_baris(kiri)
        if baris:
            info["Nama"] = teks_baris(baris[0])

    m = RE_PERIODE.search(teks)
    if not m:
        raise ValueError("Periode transaksi tidak ditemukan di PDF BNI.")
    set_periode(info, _tanggal_periode(*m.group(1, 2, 3)), _tanggal_periode(*m.group(4, 5, 6)))
    return info


# --------------------------------------------------------------------------
# Halaman
# --------------------------------------------------------------------------
def _baris_sel(page, bawah_header):
    """Batas vertikal tiap baris transaksi = kotak latar sel kolom 'Posting Date' (x0 ~ 14, lebar ~110)."""
    kotak = {}
    for r in page.rects:
        if r["x0"] < 20 and 90 < r["x1"] - r["x0"] < 140 and r["top"] >= bawah_header - 2:
            kotak[(round(r["top"]), round(r["bottom"]))] = (r["top"], r["bottom"])
    return sorted(kotak.values())


def _halaman(page, hasil, pending):
    page = page.dedupe_chars(tolerance=1)
    words = page.extract_words(x_tolerance=2)

    def kepala(teks, x_maks=None):
        c = [w for w in words if w["text"] == teks and (x_maks is None or w["x0"] < x_maks)]
        return min(c, key=lambda w: w["top"]) if c else None

    h_posting, h_journal = kepala("Posting", 100), kepala("Journal")
    h_amount, h_dbcr, h_saldo = kepala("Amount"), kepala("DB/CR"), kepala("Balance")
    if not all((h_posting, h_journal, h_amount, h_dbcr, h_saldo)):
        return                                                  # bukan halaman tabel
    bawah_header = max(w["bottom"] for w in words if abs(w["top"] - h_posting["top"]) < 6)
    x_desc = h_journal["x1"] + 1                                # keterangan mulai sesudah kolom Journal
    x_amount_kanan = h_amount["x1"]                             # angka Amount rata kanan, melewati header
    x_dbcr0, x_dbcr1 = h_dbcr["x0"] - 4, h_dbcr["x1"] + 4

    body = [w for w in words if w["top"] > bawah_header]
    baris_body = kelompok_baris(body)

    # saldo awal ("Ledger Balance:") -> dibuang dari body
    for b in baris_body:
        if teks_baris(b).startswith("Ledger Balance"):
            kand = [w for w in b if w["x0"] >= x_dbcr1]
            if kand and hasil["saldo_awal"] is None:
                v = _nilai(page, kand[-1])
                hasil["saldo_awal"] = v[0] if v else None
            ids = {id(w) for w in b}
            body = [w for w in body if id(w) not in ids]
            break

    # ringkasan di bawah tabel -> potong body
    rg = hasil["ringkasan_umum"]
    batas = None
    for b in baris_body:
        t = teks_baris(b)
        m_end = re.match(r"Ending\s+Balance\s*:\s*([\d,]+\.\d{2})", t)
        m_deb = re.match(r"Total\s+Deb[ei]t\s*:\s*(\d+)\s+([\d,]+\.\d{2})", t)
        m_kre = re.match(r"Total\s+Credit\s*:\s*(\d+)\s+([\d,]+\.\d{2})", t)
        if m_end or m_deb or m_kre:
            batas = b[0]["top"] - 1 if batas is None else min(batas, b[0]["top"] - 1)
        if m_end:
            rg["saldo_akhir"] = angka(m_end.group(1))
        if m_deb:
            rg["n_debit"], rg["total_debit"] = int(m_deb.group(1)), angka(m_deb.group(2))
        if m_kre:
            rg["n_kredit"], rg["total_kredit"] = int(m_kre.group(1)), angka(m_kre.group(2))
    if batas is not None:
        body = [w for w in body if w["top"] < batas]

    # baris transaksi
    tanggal = sorted((w for w in body if w["x0"] < 40 and RE_TGL.match(w["text"])), key=lambda w: w["top"])
    sel = _baris_sel(page, bawah_header)
    if len(sel) != len(tanggal):
        # Tanpa batas sel, keterangan multi-baris (rata tengah vertikal) bisa menempel ke baris
        # yang salah dan itu tidak terdeteksi validasi total/saldo -> lebih baik berhenti.
        raise ValueError("Batas baris tabel BNI tidak terbaca (layout PDF berbeda dari yang dikenal).")
    kelompok = [[] for _ in sel]
    lanjutan = []
    for w in body:
        if w["x0"] < x_desc and not (w["x0"] < 40 and RE_TGL.match(w["text"])):
            continue                         # kolom waktu / tanggal efektif / branch / journal
        mid = (w["top"] + w["bottom"]) / 2
        if not sel or mid < sel[0][0]:
            lanjutan.append(w)
            continue
        i = next((k for k, (a, z) in enumerate(sel) if a <= mid < z), None)
        if i is None:
            i = min(range(len(sel)), key=lambda k: min(abs(mid - sel[k][0]), abs(mid - sel[k][1])))
        kelompok[i].append(w)

    if lanjutan and pending:
        teks = " ".join(teks_baris(b) for b in kelompok_baris(
            [w for w in lanjutan if x_desc <= w["x0"] and w["x1"] <= x_amount_kanan]))
        if teks.strip():
            pending[-1]["baris"].append(teks)

    for tgl_word, ws in zip(tanggal, kelompok):
        uraian, jumlah, arah, saldo = [], None, None, None
        for w in ws:
            t = w["text"]
            if w is tgl_word or w["x0"] < x_desc:
                continue
            if w["x0"] >= x_dbcr1 - 2 and RE_ANGKA_LONGGAR.match(t):          # kolom Balance
                v = _nilai(page, w)
                saldo = v[0] if v else None
            elif x_dbcr0 <= w["x0"] and w["x1"] <= x_dbcr1 and t in ("K", "D", "C"):
                arah = "D" if t == "D" else "K"
            elif w["x1"] > x_amount_kanan and w["x1"] <= x_dbcr0 + 4 and RE_ANGKA_LONGGAR.match(t):
                jumlah = _nilai(page, w)                                       # kolom Amount
            else:
                uraian.append(w)
        if not jumlah or arah is None:
            continue                         # baris tanpa nominal/arah bukan transaksi
        pending.append({
            "tgl": datetime.strptime(tgl_word["text"], "%d/%m/%Y"),
            "baris": [teks_baris(b) for b in kelompok_baris(uraian)],
            "kand": jumlah, "arah": arah, "saldo": saldo,
        })


# --------------------------------------------------------------------------
# Keterangan
# --------------------------------------------------------------------------
def _segmen(uraian):
    seg = [s.strip() for s in re.split(r"\s*\|\s*", re.sub(r"\s+", " ", uraian)) if s.strip()]
    return [s for s in seg if not re.fullmatch(r"0+", s)]       # buang segmen referensi 0000000000000000


def _lari_kapital(teks):
    """Awal teks yang seluruh katanya berhuruf kapital/angka (nama), berhenti di kata ber-huruf kecil."""
    out = []
    for tok in teks.split():
        if re.search(r"[a-z]", tok):
            break
        out.append(tok)
    return " ".join(out) or None


def _tanpa_honorifik(teks):
    tok = teks.split()
    if tok and tok[0].upper().rstrip(".") in HONORIFIK:
        tok = tok[1:]
    return " ".join(tok) or None


def nama_lawan(segmen):
    """Nama lawan transaksi dari segmen keterangan (segmen[0] = jenis transaksi)."""
    if len(segmen) < 2:
        return None
    jenis = segmen[0].upper()
    if jenis.startswith("TRANSFER DARI") or jenis.startswith("TRANSFER KE"):
        m = RE_REK_PINDAH.match(segmen[1])
        return _tanpa_honorifik(m.group(2)) if m else None
    if jenis.startswith("SETOR TUNAI"):
        return segmen[1]
    if jenis.startswith("KREDIT LAIN"):
        return _lari_kapital(re.sub(r"^\d{2,4}\s+", "", segmen[1]))
    if jenis.startswith("TRF/PAY/TOP-UP"):
        s = segmen[-1]
        if len(segmen) < 3 or re.match(r"^\d{8,}\b", s):         # tujuan hanya nomor rekening, tanpa nama
            return None
        if len(s) > PANJANG_NAMA_ECHANNEL and s[PANJANG_NAMA_ECHANNEL] == " ":
            s = s[:PANJANG_NAMA_ECHANNEL]
        return _lari_kapital(s)
    return None


def _buat(p):
    seg = _segmen(" ".join(p["baris"]))
    jenis = seg[0] if seg else ""
    ket_lengkap = " | ".join(seg[1:]) or None
    keluar = p["jumlah"] if p["arah"] == "D" else None
    masuk = p["jumlah"] if p["arah"] == "K" else None
    return buat_transaksi(p["tgl"], jenis, ket_lengkap, nama_lawan(seg), keluar, masuk,
                          p["saldo"], detail=seg)


# --------------------------------------------------------------------------
# Pilih angka yang tepat bila ada karakter nyasar, lalu bentuk transaksi
# --------------------------------------------------------------------------
def _selesaikan(pending, hasil):
    saldo_akhir = hasil["ringkasan_umum"].get("saldo_akhir")
    sebelum = hasil["saldo_awal"]
    dikoreksi = gagal = 0
    for i, p in enumerate(pending):
        s = p["saldo"]
        if s is None and i == len(pending) - 1:
            s = saldo_akhir                                    # baris terakhir tidak mencetak saldo
        pilihan = p["kand"]
        if len(pilihan) > 1:
            cocok_ = [c for c in pilihan
                      if sebelum is not None and s is not None
                      and abs(abs(s - sebelum) - c) <= TOLERANSI]
            if len(cocok_) == 1:
                pilihan, dikoreksi = cocok_, dikoreksi + 1
            else:
                gagal += 1
        p["jumlah"] = pilihan[0]
        if s is not None:
            sebelum = s
    if dikoreksi:
        hasil["catatan"].append(f"{dikoreksi} angka mutasi tercetak ganda di PDF; nilai yang benar "
                                "dipilih berdasarkan selisih saldo.")
    if gagal:
        hasil["catatan"].append(f"{gagal} angka mutasi ambigu dan tidak bisa dipastikan dari saldo; "
                                "periksa baris tersebut.")

    hasil["transaksi"] = [_buat(p) for p in pending]
    # saldo baris terakhir = Ending Balance bila saldo berjalan cocok
    trx = hasil["transaksi"]
    if trx and trx[-1]["Saldo"] is None and saldo_akhir is not None:
        pra = next((t["Saldo"] for t in reversed(trx[:-1]) if t["Saldo"] is not None), hasil["saldo_awal"])
        if pra is not None:
            ujung = pra + (trx[-1]["Kredit"] or 0) - (trx[-1]["Debit"] or 0)
            if len(trx) >= 1 and abs(ujung - saldo_akhir) <= TOLERANSI:
                trx[-1]["Saldo"] = saldo_akhir


def proses(file_obj):
    hasil = {"transaksi": [], "saldo_awal": None, "ringkasan": {}, "ringkasan_umum": {}, "catatan": []}
    pending = []
    with pdfplumber.open(file_obj) as pdf:
        info = _header(pdf.pages[0].dedupe_chars(tolerance=1))
        for page in pdf.pages:
            _halaman(page, hasil, pending)
    if hasil["saldo_awal"] is not None:
        hasil["ringkasan_umum"]["saldo_awal"] = hasil["saldo_awal"]
    _selesaikan(pending, hasil)
    return info, hasil
