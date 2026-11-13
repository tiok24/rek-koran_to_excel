"""
Pintu masuk tunggal untuk semua format:  PDF BCA, CIMB, Danamon, BRI  dan  CSV.

    bank, info, hasil = proses_pdf_auto(file_atau_bytes)            # deteksi bank otomatis
    bank, info, hasil = proses_pdf_auto(file_atau_bytes, bank="BRI")  # paksa bank tertentu
    ok, pesan, ringkas = periksa(bank, hasil)

`info` dan `hasil` berskema sama untuk semua bank, jadi pdf_ke_rekor_bca.tulis_excel()
dapat dipakai langsung (lewat parameter bank=...).
"""

import io

import pdfplumber

import pdf_ke_rekor_bca as bca
from . import bca_csv, bri, cimb, csv_parser, danamon
from .common import nama_file_hasil, validasi_umum  # noqa: F401

PARSER_PDF = {"CIMB": cimb, "DANAMON": danamon, "BRI": bri}
BANK_PDF = ["BCA", "CIMB", "DANAMON", "BRI"]


def _ke_stream(sumber):
    if isinstance(sumber, (bytes, bytearray)):
        return io.BytesIO(sumber)
    if hasattr(sumber, "seek"):
        sumber.seek(0)
    return sumber


def _teks_halaman_1(sumber):
    try:
        with pdfplumber.open(_ke_stream(sumber)) as pdf:
            return pdf.pages[0].extract_text() or ""
    except Exception as e:  # noqa: BLE001
        raise ValueError("File bukan PDF yang valid, rusak, atau diproteksi password "
                         f"({type(e).__name__}).") from e


def _mirip_bca(teks):
    t = teks.upper()
    return "MUTASI CR" in t or ("NO. REKENING" in t and "MATA UANG" in t and "PERIODE" in t) \
        or "BCA" in t


def deteksi_bank(sumber):
    """'BCA' | 'CIMB' | 'DANAMON' | 'BRI' | None. Urutan penting: teks transaksi BRI bisa
    menyebut bank lain (mis. 'DANAMON-BRI'), jadi ciri judul dokumen diperiksa dulu."""
    teks = _teks_halaman_1(sumber)
    for nama in ("BRI", "CIMB", "DANAMON"):
        if PARSER_PDF[nama].cocok(teks):
            return nama
    if _mirip_bca(teks):
        return "BCA"
    return None


def _normalisasi_bca(info, hasil):
    """Lengkapi hasil BCA dengan kolom siap pakai & ringkasan umum (perilaku tulis_excel tidak berubah)."""
    for t in hasil["transaksi"]:
        t["ket_bank"] = bca.keterangan_bank(t["Jenis Transaksi"], t["jenis_extra"])
        t["ket_lengkap"] = bca.keterangan_lengkap(t["detail"])
        t["nama_lawan"] = bca.nama_lawan(t["Jenis Transaksi"], t["detail"])
    bln, thn = bca.bulan_tahun(info["Periode"])
    if bln:
        from datetime import datetime
        info["periode_awal"] = info["periode_akhir"] = datetime(thn, bln, 1)
    return info, hasil


def proses_pdf_auto(sumber, bank=None):
    """-> (bank, info, hasil). `sumber` = bytes / file-like / path."""
    bank = bank or deteksi_bank(sumber)
    if bank is None:
        raise ValueError("Bank tidak dikenali. Format PDF yang didukung: "
                         + ", ".join(BANK_PDF) + ". Pilih bank secara manual bila perlu.")
    try:
        if bank == "BCA":
            info, hasil = bca.proses_pdf(_ke_stream(sumber))
            info, hasil = _normalisasi_bca(info, hasil)
        else:
            info, hasil = PARSER_PDF[bank].proses(_ke_stream(sumber))
    except Exception as e:  # noqa: BLE001
        raise ValueError(f"Gagal membaca PDF sebagai format {bank}: {str(e).rstrip('.')}. "
                         "Bila bank salah terdeteksi/dipilih, ubah pilihan bank.") from e
    if not hasil["transaksi"]:
        raise ValueError(f"Tidak ada transaksi yang terbaca (bank terdeteksi: {bank}). "
                         "PDF mungkin hasil scan atau formatnya berbeda.")
    return bank, info, hasil


def periksa(bank, hasil):
    """-> (ok, daftar_pesan, ringkasan). BCA memakai cek_laporan() asli."""
    if bank == "BCA":
        v = bca.validasi(hasil)
        ok = bca.cek_laporan(hasil)
        trx = hasil["transaksi"]
        ringkas = {
            "saldo_awal": hasil["saldo_awal"],
            "saldo_akhir": v["saldo_akhir_pdf"],
            "total_masuk": sum(t["Kredit"] or 0 for t in trx),
            "total_keluar": sum(t["Debit"] or 0 for t in trx),
        }
        pesan = [] if ok else ["Total/jumlah transaksi tidak cocok dengan ringkasan di PDF."]
        return ok, pesan, ringkas
    v = validasi_umum(hasil)
    return v["ok"], v["pesan"], v


def nama_file(bank, info):
    if bank == "BCA":
        bln, thn = bca.bulan_tahun(info["Periode"])
        return f"REKOR_{info['No Rekening']}_{thn}-{bln:02d}.xlsx"
    return nama_file_hasil(info)
