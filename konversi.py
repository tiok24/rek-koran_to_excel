"""
Konversi mutasi rekening -> Excel REKOR lewat baris perintah (PDF BCA/CIMB/Danamon/BRI dan CSV).

    python konversi.py a.pdf b.pdf data.csv        # file tertentu
    python konversi.py folder_mutasi               # semua .pdf/.csv di folder
    python konversi.py -t REKOR_2026.xlsx -o hasil_excel *.pdf

Opsi PDF : --bank BRI            paksa bank (bawaan: deteksi otomatis)
Opsi CSV : --nama "PT CONTOH" --rekening 1234567890 --label-bank MANDIRI
           --format-angka id|en|auto   --tukar-arah   --saldo-awal 1000000
           (kolom CSV ditebak otomatis dari nama header)
"""

import argparse
import sys
from pathlib import Path

from bank_parsers import BANK_PDF, csv_parser
from bank_parsers.pipeline import olah_csv, olah_pdf


def kumpulkan(argumen):
    out = []
    for p in map(Path, argumen):
        if p.is_dir():
            out += sorted(x for x in p.iterdir() if x.suffix.lower() in (".pdf", ".csv"))
        elif p.suffix.lower() in (".pdf", ".csv") and p.exists():
            out.append(p)
        else:
            print(f"[LEWATI] {p}: bukan file .pdf/.csv/folder yang valid")
    return out


def main():
    ap = argparse.ArgumentParser(description="Mutasi rekening (PDF/CSV) -> Excel template REKOR")
    ap.add_argument("berkas", nargs="+", help="file .pdf / .csv atau folder")
    ap.add_argument("-t", "--template", default="REKOR_2026.xlsx")
    ap.add_argument("-o", "--output", default="hasil_excel")
    ap.add_argument("--bank", choices=BANK_PDF, help="paksa bank untuk PDF")
    ap.add_argument("--nama", default="")
    ap.add_argument("--rekening", default="")
    ap.add_argument("--label-bank", default="BANK")
    ap.add_argument("--format-angka", choices=["auto", "id", "en"], default="auto")
    ap.add_argument("--tukar-arah", action="store_true")
    ap.add_argument("--saldo-awal", type=float, default=None)
    a = ap.parse_args()

    template = next((p for p in (Path(a.template), Path(__file__).resolve().parent / a.template)
                     if p.exists()), None)
    if template is None:
        sys.exit(f"Template '{a.template}' tidak ditemukan. Taruh di folder yang sama atau pakai -t.")
    tpl = template.read_bytes()

    daftar = kumpulkan(a.berkas)
    if not daftar:
        sys.exit("Tidak ada file yang bisa diproses.")
    keluar = Path(a.output)
    keluar.mkdir(parents=True, exist_ok=True)

    sukses = 0
    for p in daftar:
        try:
            data = p.read_bytes()
            if p.suffix.lower() == ".csv":
                header, _, _ = csv_parser.muat_csv(data)
                peta = csv_parser.tebak_pemetaan(header)
                print(f"      kolom CSV: { {k: v for k, v in peta.items() if v} }")
                r = olah_csv(data, tpl, peta, {
                    "nama": a.nama, "no_rekening": a.rekening, "label_bank": a.label_bank,
                    "format_angka": a.format_angka, "tukar_arah": a.tukar_arah,
                    "saldo_awal": a.saldo_awal})
            else:
                r = olah_pdf(data, tpl, a.bank)
            (keluar / r["nama_file"]).write_bytes(r["excel"])
            status = "OK" if r["ok"] else ("INFO" if not r["bisa_diverifikasi"] else "CEK")
            print(f"[{status}] {p.name} [{r['bank']}] -> {r['nama_file']} ({len(r['df'])} transaksi)")
            for pesan in ([] if r["ok"] else r["pesan"]):
                print(f"      {pesan}")
            sukses += 1
        except Exception as e:  # noqa: BLE001
            print(f"[ERROR] {p.name}: {e}")
    print(f"\nSelesai! {sukses} file Excel disimpan di folder '{keluar}'.")


if __name__ == "__main__":
    main()
