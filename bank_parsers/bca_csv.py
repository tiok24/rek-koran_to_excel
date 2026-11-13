"""
Parser CSV BCA (KlikBCA Bisnis): "Informasi Rekening - Rekening Koran" dan "Mutasi Rekening".

Ciri format:
    - Judul laporan + No. rekening / Nama / Periode / Kode Mata Uang di atas tabel.
    - Kolom: Tanggal Transaksi | Keterangan | Cabang | Jumlah | Saldo
    - Jumlah diakhiri penanda arah:  "54,000.00DB" atau "50,896,182.00 CR" (ada/tidak ada spasi).
    - Tanggal "dd/mm" (tanpa tahun) atau "dd/mm/yyyy"; transaksi belum posting bertanggal "PEND".
    - Pada varian "Rekening Koran", keterangan satu transaksi terpecah ke beberapa baris CSV
      (baris lanjutan diawali koma), sehingga harus disambung dulu sebelum dibaca.
    - Footer: Saldo Awal, Mutasi Debet (+jumlah transaksi), Mutasi Kredit (+jumlah), Saldo Akhir
      -> dipakai untuk memverifikasi hasil.

Konvensi sama dengan parser lain: t["Debit"] = uang KELUAR (DB), t["Kredit"] = uang MASUK (CR).
"""

import csv
import re

import pdf_ke_rekor_bca as bca
from .common import buat_transaksi, set_periode, tahun_untuk_bulan
from .csv_parser import _decode, parse_angka, parse_tanggal

NAMA = "BCA CSV"
LABEL_EXCEL = "BCA"

RE_TGL = re.compile(r"^(\d{2})/(\d{2})(?:/(\d{2}|\d{4}))?$")
RE_JUMLAH = re.compile(r"^\s*([\d,]+\.\d{2})\s*(DB|CR)\s*$", re.I)
RE_NOMINAL_DEPAN = re.compile(r"^(\d+\.\d{2})(?=\D|$)\s*")
RE_NOMINAL_AKHIR = re.compile(r"\s+(\d+\.\d{2})\s*$")
TOLERANSI = 0.01


def cocok(data):
    """True bila bytes CSV berformat KlikBCA Bisnis (rekening koran / mutasi rekening)."""
    teks = _decode(data) if isinstance(data, (bytes, bytearray)) else str(data)
    kepala = teks[:1500]
    return ("Informasi Rekening" in kepala
            and re.search(r"Tanggal Transaksi\W+Keterangan\W+Cabang\W+Jumlah\W+Saldo", teks[:2500]) is not None)


# --------------------------------------------------------------------------
def _sambung_baris(teks):
    """Baris yang diawali koma = lanjutan transaksi sebelumnya -> digabung jadi satu baris."""
    hasil = []
    for g in teks.splitlines():
        if g.startswith(",") and hasil:
            hasil[-1] += g
        elif g.strip():
            hasil.append(g)
    return hasil


def _meta(baris):
    """Nama, No. rekening, periode & footer ringkasan dari baris-baris non-transaksi."""
    m = {"nama": "", "no_rekening": "", "periode": None, "mata_uang": None,
         "saldo_awal": None, "saldo_akhir": None,
         "total_debit": None, "n_debit": None, "total_kredit": None, "n_kredit": None}
    for g in baris:
        sel = next(csv.reader([g]), [])
        if not sel:
            continue
        s = sel[0].strip()
        if re.match(r"No\.?\s*rekening", s, re.I):
            m["no_rekening"] = s.split(":", 1)[1].strip()
        elif re.match(r"Nama\s*:", s, re.I):
            m["nama"] = s.split(":", 1)[1].strip()
        elif re.match(r"Periode\s*:", s, re.I):
            p = re.findall(r"\d{1,2}/\d{1,2}/\d{2,4}", s)
            if len(p) == 2:
                m["periode"] = (parse_tanggal(p[0]), parse_tanggal(p[1]))
        elif re.match(r"Kode Mata Uang", s, re.I):
            m["mata_uang"] = s.split(":", 1)[1].strip()
        elif re.match(r"Saldo Awal\s*:", s, re.I):
            m["saldo_awal"] = parse_angka(s.split(":", 1)[1], "en")
        elif re.match(r"Saldo Akhir\s*:", s, re.I):
            m["saldo_akhir"] = parse_angka(s.split(":", 1)[1], "en")
        elif re.match(r"Mutasi Debet\s*:", s, re.I):
            m["total_debit"] = parse_angka(s.split(":", 1)[1], "en")
            m["n_debit"] = int(sel[1]) if len(sel) > 1 and sel[1].strip().isdigit() else None
        elif re.match(r"Mutasi Kredit\s*:", s, re.I):
            m["total_kredit"] = parse_angka(s.split(":", 1)[1], "en")
            m["n_kredit"] = int(sel[1]) if len(sel) > 1 and sel[1].strip().isdigit() else None
    return m


def _pecah_keterangan(bagian, jumlah):
    """Bagian-bagian kolom Keterangan -> (jenis, detail[list]).

    Varian Rekening Koran : tiap baris CSV = satu bagian (jenis di bagian pertama).
    Varian Mutasi Rekening: satu sel, bagian dipisah >=2 spasi.
    Nominal transfer yang menempel di depan/akhir teks (mis. '93350000.00LESTARI') dibuang
    selama sama dengan jumlah transaksi, karena sudah ada di kolom Debit/Kredit.
    """
    if len(bagian) == 1:
        bagian = [b for b in re.split(r"\s{2,}", bagian[0].strip()) if b.strip()]
    bagian = [re.sub(r"\s+", " ", b).strip() for b in bagian if b and b.strip()]
    if not bagian:
        return "", []

    def sama_jumlah(teks):
        return abs(float(teks) - jumlah) < TOLERANSI

    jenis, detail = bagian[0], bagian[1:]
    m = RE_NOMINAL_AKHIR.search(jenis)
    if m and sama_jumlah(m.group(1)):
        jenis = jenis[:m.start()].strip()
    for i, d in enumerate(detail):
        m = RE_NOMINAL_DEPAN.match(d)
        if m and sama_jumlah(m.group(1)):
            detail[i] = d[m.end():].strip()
            break
    detail = [d for d in detail if d]
    return jenis, detail


def proses_csv(data):
    """bytes CSV KlikBCA Bisnis -> (info, hasil), skema sama dengan parser PDF."""
    teks = _decode(data) if isinstance(data, (bytes, bytearray)) else str(data)
    baris = _sambung_baris(teks)
    meta = _meta(baris)
    if not meta["periode"]:
        raise ValueError("Baris 'Periode' tidak ditemukan di CSV BCA.")
    awal, akhir = meta["periode"]

    transaksi, catatan, dilewati = [], [], 0
    for g in baris:
        sel = next(csv.reader([g]), [])
        if len(sel) < 5:
            continue
        kode = sel[0].strip()
        pending = kode.upper() == "PEND"
        mt = RE_TGL.match(kode)
        if not (pending or mt):
            continue

        mj = RE_JUMLAH.match(sel[-2])
        saldo = parse_angka(sel[-1], "en")
        if not mj:
            dilewati += 1
            catatan.append(f"Baris dilewati (jumlah tidak terbaca): {kode} {sel[1][:40]}")
            continue
        jumlah = float(mj.group(1).replace(",", ""))
        arah = mj.group(2).upper()

        if pending:
            tgl = akhir
        else:
            d, bln, thn = int(mt.group(1)), int(mt.group(2)), mt.group(3)
            thn = (int(thn) + (2000 if len(thn) == 2 else 0)) if thn else tahun_untuk_bulan(bln, awal, akhir)
            tgl = parse_tanggal(f"{d:02d}/{bln:02d}/{thn}")
            if tgl is None:
                dilewati += 1
                catatan.append(f"Baris dilewati (tanggal tidak valid): {kode}")
                continue

        jenis, detail = _pecah_keterangan(sel[1:-3] if len(sel) > 5 else sel[1:2], jumlah)
        t = buat_transaksi(
            tgl, bca.keterangan_bank(jenis, []), bca.keterangan_lengkap(detail),
            _bersihkan_nama(bca.nama_lawan(jenis, detail)),
            jumlah if arah == "DB" else None, jumlah if arah == "CR" else None,
            saldo, detail=detail)
        t["pending"] = pending
        transaksi.append(t)

    if not transaksi:
        raise ValueError("Tidak ada transaksi yang terbaca dari CSV BCA.")

    n_pend = sum(1 for t in transaksi if t["pending"])
    if n_pend:
        catatan.insert(0, f"{n_pend} transaksi berstatus PEND (belum posting, tanpa tanggal di CSV) "
                          f"diberi tanggal akhir periode ({akhir:%d/%m/%Y}).")

    tgl_ada = [t["Tanggal"] for t in transaksi if not t["pending"]] or [akhir]
    info = {"Cabang": None, "Nama": meta["nama"], "No Rekening": meta["no_rekening"] or "BCA",
            "Periode": None, "Mata Uang": meta["mata_uang"]}
    set_periode(info, min(tgl_ada), max(tgl_ada))

    ringkasan = {k: meta[k] for k in ("saldo_awal", "saldo_akhir", "total_debit", "total_kredit",
                                      "n_debit", "n_kredit")}
    saldo_awal = meta["saldo_awal"]
    if saldo_awal is None:                       # fallback: hitung dari baris pertama
        t0 = transaksi[0]
        saldo_awal = (t0["Saldo"] - (t0["Kredit"] or 0) + (t0["Debit"] or 0)
                      if t0["Saldo"] is not None else 0.0)
    hasil = {"transaksi": transaksi, "saldo_awal": saldo_awal, "ringkasan": {},
             "ringkasan_umum": ringkasan, "dilewati": dilewati, "urutan_dibalik": False,
             "catatan": catatan}
    return info, hasil


def _bersihkan_nama(nama):
    return re.sub(r"\s+", " ", nama).strip(" ,") or None if nama else nama
