"""
Lapisan tipis di atas parser: bytes berkas -> hasil siap tampil (Excel di memori, tabel
pratinjau, status verifikasi). Dipakai bersama oleh app.py (Streamlit) dan konversi.py (CLI).
"""

import io

import pandas as pd

from pdf_ke_rekor_bca import tulis_excel
from . import csv_parser, nama_file, periksa, proses_pdf_auto


def _susun(bank, info, hasil, template_bytes, label=None):
    ok, pesan, ringkas = periksa(bank, hasil)
    buf = io.BytesIO()
    tulis_excel(io.BytesIO(template_bytes), info, hasil, buf, bank=label or bank)

    trx = hasil["transaksi"]
    df = pd.DataFrame({
        "No": range(1, len(trx) + 1),
        "Tanggal": [t["Tanggal"].date() for t in trx],
        "Keterangan Bank": [t["ket_bank"] for t in trx],
        "Keterangan 1 (Lengkap)": [t["ket_lengkap"] for t in trx],
        "Nama Lawan Transaksi": [t["nama_lawan"] for t in trx],
        "Debit (Masuk)": [t["Kredit"] for t in trx],
        "Kredit (Keluar)": [t["Debit"] for t in trx],
        "Saldo Bank": [t["Saldo"] for t in trx],
    })
    return {
        "bank": bank,
        "info": info,
        "nama_file": nama_file(bank, info),
        "excel": buf.getvalue(),
        "ok": ok,
        "pesan": pesan,
        # hanya CSV tanpa saldo yang tidak punya pembanding sama sekali
        "bisa_diverifikasi": ringkas.get("bisa_diverifikasi", True) if isinstance(ringkas, dict) else True,
        "saldo_awal": ringkas.get("saldo_awal", hasil["saldo_awal"]),
        "saldo_akhir": ringkas.get("saldo_akhir"),
        "total_masuk": df["Debit (Masuk)"].sum(),
        "total_keluar": df["Kredit (Keluar)"].sum(),
        "df": df,
        "dilewati": hasil.get("dilewati", 0),
        "urutan_dibalik": hasil.get("urutan_dibalik", False),
    }


def olah_pdf(data: bytes, template_bytes: bytes, bank=None):
    """PDF (BCA/CIMB/Danamon/BRI) -> hasil. bank=None = deteksi otomatis."""
    bank, info, hasil = proses_pdf_auto(data, bank=bank)
    return _susun(bank, info, hasil, template_bytes)


def olah_csv(data: bytes, template_bytes: bytes, pemetaan: dict, opsi: dict):
    """CSV -> hasil. opsi: nama, no_rekening, label_bank, format_angka, tukar_arah,
    saldo_awal (float|None), hari_dulu, pemisah (None = otomatis)."""
    header, baris, _ = csv_parser.muat_csv(data, opsi.get("pemisah"))
    info, hasil = csv_parser.konversi_csv(
        header, baris, pemetaan,
        nama=opsi.get("nama", ""), no_rekening=opsi.get("no_rekening", ""),
        format_angka=opsi.get("format_angka", "auto"), tukar_arah=opsi.get("tukar_arah", False),
        saldo_awal=opsi.get("saldo_awal"), hari_dulu=opsi.get("hari_dulu", True))
    r = _susun("CSV", info, hasil, template_bytes, label=opsi.get("label_bank") or "BANK")
    return r
