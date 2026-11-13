"""
PDF mutasi rekening giro BCA  ->  Excel (format template REKOR).
SATU FILE PDF = SATU FILE EXCEL.

Alur:
  1. Pilih file PDF yang ingin diekstrak (jendela pilih file, atau tulis di perintah).
  2. Setiap PDF dibaca dan transaksinya diekstrak.
  3. Hasilnya dimasukkan ke salinan template Excel (isi contoh di template dikosongkan,
     format/border/lebar kolom tetap).
  4. Setiap PDF menghasilkan 1 file Excel.

Instalasi:
    pip install pdfplumber openpyxl

Pemakaian:
    python pdf_ke_rekor_bca.py                       # buka jendela pilih file PDF
    python pdf_ke_rekor_bca.py a.pdf b.pdf           # file tertentu
    python pdf_ke_rekor_bca.py folder_pdf            # semua PDF di folder
    Opsi:  -t REKOR_2026.xlsx   (template, bawaan REKOR_2026.xlsx)
           -o hasil_excel       (folder hasil, bawaan hasil_excel)

Nama file hasil: REKOR_<no rekening>_<tahun>-<bulan>.xlsx

Isi kolom template:
  C  Tanggal              : tanggal transaksi
  D  Keterangan Bank      : jenis transaksi di PDF (mis. TRSF E-BANKING CR), tanpa tanggal
  F  Keterangan 2         : nama lawan transaksi (lihat fungsi nama_lawan)
  I  Debit                : uang MASUK  (= mutasi CR di PDF)
  J  Kredit               : uang KELUAR (= mutasi DB di PDF)
  K  Saldo                : rumus =K(atas)+I-J ; K7 = Saldo Awal dari PDF
  L  Saldo Bank           : saldo yang tercetak di PDF (hanya baris yang ada saldonya)
  M  Selisih              : =K-L  (hanya di baris yang L-nya terisi)
  E  Keterangan 1         : keterangan transaksi lengkap dari PDF (semua baris digabung)
  G, H dibiarkan kosong (diisi manual).
"""

import argparse
import hashlib
import re
import sys
from bisect import bisect_right
from copy import copy
from datetime import datetime
from pathlib import Path

import pdfplumber
from openpyxl import load_workbook

# Batas kolom (koordinat X dalam poin, halaman A4 lebar 595).
X_JENIS = 80      # kolom Jenis Transaksi mulai
X_DETAIL = 190    # kolom Keterangan (detail) mulai
X_CBG = 300       # kolom CBG mulai
X_MUTASI = 350    # kolom Mutasi mulai
X_SALDO = 480     # kolom Saldo mulai

BARIS_SALDO_AWAL = 7
BARIS_DATA_PERTAMA = 8
JML_KOLOM = 13  # A..M
BANK = "BCA"
TOLERANSI = 0.01

# --------------------------------------------------------------------------
# BAGIAN 1: BACA PDF
# --------------------------------------------------------------------------
BULAN = {
    "JANUARI": 1, "FEBRUARI": 2, "MARET": 3, "APRIL": 4, "MEI": 5, "JUNI": 6,
    "JULI": 7, "AGUSTUS": 8, "SEPTEMBER": 9, "OKTOBER": 10, "NOVEMBER": 11,
    "DESEMBER": 12,
}
RE_TGL = re.compile(r"^\d{2}/\d{2}$")
RE_ANGKA = re.compile(r"^\d{1,3}(?:,\d{3})*\.\d{2}$")


def angka(teks):
    return float(teks.replace(",", ""))


def kelompok_baris(words, tol=2.5):
    """Kelompokkan kata menjadi baris berdasarkan posisi vertikal."""
    baris, cur, cur_top = [], [], None
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if cur_top is None or w["top"] - cur_top > tol:
            if cur:
                baris.append(cur)
            cur, cur_top = [w], w["top"]
        else:
            cur.append(w)
    if cur:
        baris.append(cur)
    return [sorted(b, key=lambda w: w["x0"]) for b in baris]


def teks_baris(b):
    return " ".join(w["text"] for w in b)


# --------------------------------------------------------------------------
# HEADER HALAMAN 1
# --------------------------------------------------------------------------
def baca_header(page):
    words = page.extract_words(x_tolerance=2)
    info = {"Cabang": None, "Nama": None, "No Rekening": None,
            "Periode": None, "Mata Uang": None}

    kiri = [w for w in words if w["x0"] < 290 and 55 < w["top"] < 170]
    baris_kiri = [teks_baris(b) for b in kelompok_baris(kiri)]
    if baris_kiri:
        info["Cabang"] = baris_kiri[0]
    if len(baris_kiri) > 1:
        info["Nama"] = baris_kiri[1]

    kanan = " ".join(teks_baris(b) for b in kelompok_baris(
        [w for w in words if w["x0"] >= 290 and 55 < w["top"] < 170]))
    for kunci, pola in (
        ("No Rekening", r"NO\.?\s*REKENING\s*:\s*(\d+)"),
        ("Periode", r"PERIODE\s*:\s*([A-Z]+\s+\d{4})"),
        ("Mata Uang", r"MATA UANG\s*:\s*([A-Z]{3})"),
    ):
        m = re.search(pola, kanan)
        info[kunci] = m.group(1).strip() if m else None
    return info


def bulan_tahun(periode):
    if periode:
        m = re.match(r"([A-Z]+)\s+(\d{4})", periode)
        if m and m.group(1) in BULAN:
            return BULAN[m.group(1)], int(m.group(2))
    return None, None


# --------------------------------------------------------------------------
# PARSE HALAMAN
# --------------------------------------------------------------------------
def parse_halaman(page, hasil):
    """Tambahkan transaksi dari satu halaman ke dict `hasil`."""
    words = page.extract_words(x_tolerance=2)
    header = [w for w in words if w["text"] == "TANGGAL" and w["x0"] < 100]
    if not header:
        return
    atas = header[0]["bottom"] + 2
    body = [w for w in words if w["top"] > atas]

    # potong di "Bersambung..." atau di blok ringkasan akhir
    batas = None
    for b in kelompok_baris(body):
        t = teks_baris(b)
        if t.startswith("Bersambung"):
            batas = b[0]["top"] - 1
            break
        m = re.match(r"^(SALDO AWAL|SALDO AKHIR|MUTASI CR|MUTASI DB)\s*:\s*(.*)$", t)
        if m:
            batas = batas if batas is not None else b[0]["top"] - 1
            nilai = re.findall(r"[\d,]+\.\d{2}|\d+$", m.group(2))
            hasil["ringkasan"][m.group(1)] = nilai
    if batas is not None:
        body = [w for w in body if w["top"] < batas]

    tanggal = sorted(
        (w for w in body if w["x0"] < X_JENIS and RE_TGL.match(w["text"])),
        key=lambda w: w["top"])
    starts = [w["top"] - 3 for w in tanggal]

    kelompok = [[] for _ in tanggal]
    sisa_awal = []  # kata sebelum tanggal pertama (lanjutan dari halaman lalu)
    for w in body:
        if w["x0"] < X_JENIS and not RE_TGL.match(w["text"]):
            continue
        i = bisect_right(starts, w["top"]) - 1
        (kelompok[i] if i >= 0 else sisa_awal).append(w)

    if sisa_awal and hasil["transaksi"]:
        tambahan = " ".join(teks_baris(b) for b in kelompok_baris(
            [w for w in sisa_awal if X_DETAIL <= w["x0"] < X_CBG]))
        if tambahan.strip():
            hasil["transaksi"][-1]["Keterangan"] += " | " + tambahan
            hasil["transaksi"][-1]["detail"].append(tambahan)

    for tgl_word, ws in zip(tanggal, kelompok):
        jenis = [w for w in ws if X_JENIS <= w["x0"] < X_DETAIL]
        detail = [w for w in ws if X_DETAIL <= w["x0"] < X_CBG]
        cbg = [w for w in ws if X_CBG <= w["x0"] < X_MUTASI]
        mutasi = [w for w in ws if X_MUTASI <= w["x0"] < X_SALDO]
        saldo = [w for w in ws if w["x0"] >= X_SALDO and RE_ANGKA.match(w["text"])]

        baris_jenis = [teks_baris(b) for b in kelompok_baris(jenis)]
        jenis_utama = baris_jenis[0] if baris_jenis else ""
        jenis_extra = baris_jenis[1:]
        baris_detail = [teks_baris(b) for b in kelompok_baris(detail)]
        baris_detail = [t for t in baris_detail if t.strip() not in ("", "-")]
        keterangan = " | ".join(jenis_extra + baris_detail)

        nilai_saldo = angka(saldo[0]["text"]) if saldo else None

        if jenis_utama == "SALDO AWAL":
            hasil["saldo_awal"] = nilai_saldo
            continue

        nominal = [w for w in mutasi if RE_ANGKA.match(w["text"])]
        if not nominal:
            continue
        is_debit = any(w["text"] == "DB" for w in mutasi)
        nom = angka(nominal[0]["text"])

        hasil["transaksi"].append({
            "tgl_raw": tgl_word["text"],
            "Jenis Transaksi": jenis_utama,
            "Keterangan": keterangan,
            "jenis_extra": jenis_extra,   # baris tambahan di kolom jenis (mis. "TANGGAL :12/02")
            "detail": baris_detail,       # baris-baris kolom keterangan (tanpa "-")
            "CBG": " ".join(w["text"] for w in cbg) or None,
            "Debit": nom if is_debit else None,
            "Kredit": None if is_debit else nom,
            "Saldo": nilai_saldo,
        })


def proses_pdf(path):
    hasil = {"transaksi": [], "saldo_awal": None, "ringkasan": {}}
    with pdfplumber.open(path) as pdf:
        info = baca_header(pdf.pages[0])
        for page in pdf.pages:
            parse_halaman(page, hasil)

    bln, thn = bulan_tahun(info["Periode"])
    for t in hasil["transaksi"]:
        d, m = map(int, t["tgl_raw"].split("/"))
        y = thn or datetime.now().year
        if bln == 1 and m == 12:
            y -= 1
        t["Tanggal"] = datetime(y, m, d)
    return info, hasil


# --------------------------------------------------------------------------
# VALIDASI TERHADAP RINGKASAN DI PDF
# --------------------------------------------------------------------------
def validasi(hasil):
    trx = hasil["transaksi"]
    rg = hasil["ringkasan"]
    out = {
        "saldo_awal_pdf": hasil["saldo_awal"],
        "tot_kredit_pdf": None, "n_kredit_pdf": None,
        "tot_debit_pdf": None, "n_debit_pdf": None,
        "saldo_akhir_pdf": None, "baris_saldo_tak_cocok": 0,
    }
    if "MUTASI CR" in rg and len(rg["MUTASI CR"]) >= 2:
        out["tot_kredit_pdf"], out["n_kredit_pdf"] = angka(rg["MUTASI CR"][0]), int(rg["MUTASI CR"][1])
    if "MUTASI DB" in rg and len(rg["MUTASI DB"]) >= 2:
        out["tot_debit_pdf"], out["n_debit_pdf"] = angka(rg["MUTASI DB"][0]), int(rg["MUTASI DB"][1])
    if "SALDO AKHIR" in rg and rg["SALDO AKHIR"]:
        out["saldo_akhir_pdf"] = angka(rg["SALDO AKHIR"][0])

    # cek saldo berjalan pada setiap baris yang menampilkan saldo
    if hasil["saldo_awal"] is not None:
        berjalan = hasil["saldo_awal"]
        for t in trx:
            berjalan += (t["Kredit"] or 0) - (t["Debit"] or 0)
            if t["Saldo"] is not None and abs(berjalan - t["Saldo"]) > 0.01:
                out["baris_saldo_tak_cocok"] += 1
                berjalan = t["Saldo"]  # sinkronkan agar tidak berantai
    return out


# --------------------------------------------------------------------------
# BAGIAN 2: ISI TEMPLATE EXCEL
# --------------------------------------------------------------------------
def nama_lawan(jenis, detail):
    """Isi kolom 'Keterangan 2' dari baris-baris keterangan di PDF."""
    d = [x for x in detail if x.strip()]
    if not d:
        return None
    j = jenis.upper()
    if "/FTFVA/" in d[0] and len(d) > 1:              # bayar virtual account (BPJS, dll)
        return d[1]
    if j.startswith("KR OTOMATIS") and len(d) > 1:    # kliring masuk: LLG-xxx, lalu nama
        return " ".join(d[1:])
    if j.startswith(("TRSF E-BANKING", "BI-FAST", "DB OTOMATIS", "BA JASA")):
        if len(d) > 1 and d[-1].upper() == "KBB":
            d = d[:-1]
        return d[-1]
    return " ".join(d)                                # setoran, tarikan, cek, dll


def keterangan_bank(jenis, jenis_extra):
    """Jenis transaksi untuk kolom D. Baris tambahan yang berisi tanggal
    (mis. 'TANGGAL :12/02') tidak dimasukkan."""
    sisa = [x for x in jenis_extra if not re.search(r"\d{1,2}\s*/\s*\d{1,2}", x)]
    return " ".join([jenis] + sisa).strip()


def keterangan_lengkap(detail):
    """Semua baris keterangan transaksi di PDF digabung jadi satu teks (kolom E)."""
    teks = " ".join(x.strip() for x in detail if x.strip())
    return teks or None


def cek_laporan(hasil):
    """True bila total & jumlah transaksi cocok dengan ringkasan di PDF."""
    v = validasi(hasil)
    trx = hasil["transaksi"]
    n_cr = sum(1 for t in trx if t["Kredit"] is not None)
    n_db = sum(1 for t in trx if t["Debit"] is not None)
    tot_cr = sum(t["Kredit"] or 0 for t in trx)
    tot_db = sum(t["Debit"] or 0 for t in trx)
    return (v["tot_kredit_pdf"] is not None
            and v["n_kredit_pdf"] == n_cr and v["n_debit_pdf"] == n_db
            and abs(tot_cr - v["tot_kredit_pdf"]) < TOLERANSI
            and abs(tot_db - v["tot_debit_pdf"]) < TOLERANSI
            and v["baris_saldo_tak_cocok"] == 0)


def tulis_excel(template, info, hasil, path_output, bank=BANK):
    """`bank` = label bank di sel A2 / nama sheet / kolom B. Transaksi dari parser bank
    lain boleh membawa kolom siap pakai: ket_bank, ket_lengkap, nama_lawan."""
    wb = load_workbook(template)
    ws = wb.worksheets[0]
    gaya = [copy(ws.cell(BARIS_DATA_PERTAMA, c)._style) for c in range(1, JML_KOLOM + 1)]

    # kosongkan isi contoh di template (format/border tetap)
    for r in range(BARIS_DATA_PERTAMA, ws.max_row + 1):
        for c in range(1, JML_KOLOM + 1):
            ws.cell(r, c).value = None

    bln, thn = bulan_tahun(info["Periode"])
    ws["A1"] = info["Nama"]
    ws["A2"] = f"REKENING BANK {bank} {info['No Rekening']}"
    ws["A3"] = f"TAHUN {thn}"
    ws.title = f"{bank} {info['No Rekening'][-4:]}"
    ws.cell(BARIS_SALDO_AWAL, 11).value = hasil["saldo_awal"]

    r = BARIS_SALDO_AWAL
    for no, t in enumerate(hasil["transaksi"], start=1):
        r += 1
        if "ket_bank" in t:   # parser bank lain sudah menyiapkan ketiga kolom ini
            jenis, ket1, ket2 = t["ket_bank"], t["ket_lengkap"], t["nama_lawan"]
        else:                 # BCA (perilaku asli)
            jenis = keterangan_bank(t["Jenis Transaksi"], t["jenis_extra"])
            ket1 = keterangan_lengkap(t["detail"])
            ket2 = nama_lawan(t["Jenis Transaksi"], t["detail"])
        nilai = {
            1: no, 2: bank, 3: t["Tanggal"], 4: jenis,
            5: ket1,
            6: ket2,
            9: t["Kredit"],   # uang masuk  -> kolom Debit
            10: t["Debit"],   # uang keluar -> kolom Kredit
            11: f"=K{r-1}+I{r}-J{r}",
            12: t["Saldo"],
            13: f"=K{r}-L{r}" if t["Saldo"] is not None else None,
        }
        for c in range(1, JML_KOLOM + 1):
            sel = ws.cell(r, c)
            sel._style = copy(gaya[c - 1])
            sel.value = nilai.get(c)
    wb.save(path_output)
    return r - BARIS_SALDO_AWAL


# --------------------------------------------------------------------------
# BAGIAN 3: INPUT FILE & PROGRAM UTAMA
# --------------------------------------------------------------------------
def pilih_pdf(argumen):
    """Kembalikan daftar file PDF dari argumen, atau dari jendela pilih file."""
    if argumen:
        hasil = []
        for p in map(Path, argumen):
            if p.is_dir():
                hasil += sorted(p.glob("*.pdf"))
            elif p.suffix.lower() == ".pdf" and p.exists():
                hasil.append(p)
            else:
                print(f"[LEWATI] {p}: bukan file PDF/folder yang valid")
        return hasil
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        pilihan = filedialog.askopenfilenames(
            title="Pilih file PDF mutasi rekening", filetypes=[("File PDF", "*.pdf")])
        root.destroy()
        return [Path(p) for p in pilihan]
    except Exception:  # noqa: BLE001  (tidak ada layar/tkinter)
        folder = Path("mutasi_pdf")
        print("Jendela pilih file tidak tersedia; memakai folder 'mutasi_pdf'.")
        return sorted(folder.glob("*.pdf")) if folder.is_dir() else []


def cari_template(nama):
    for lokasi in (Path(nama), Path(__file__).resolve().parent / nama):
        if lokasi.exists():
            return lokasi
    return None


def main():
    ap = argparse.ArgumentParser(description="PDF mutasi giro BCA -> Excel template REKOR (1 PDF = 1 Excel)")
    ap.add_argument("pdf", nargs="*", help="file PDF atau folder berisi PDF")
    ap.add_argument("-t", "--template", default="REKOR_2026.xlsx")
    ap.add_argument("-o", "--output", default="hasil_excel")
    a = ap.parse_args()

    template = cari_template(a.template)
    if template is None:
        sys.exit(f"Template '{a.template}' tidak ditemukan. Taruh di folder yang sama atau pakai opsi -t.")

    daftar = pilih_pdf(a.pdf)
    if not daftar:
        sys.exit("Tidak ada file PDF yang dipilih.")
    keluar = Path(a.output)
    keluar.mkdir(parents=True, exist_ok=True)

    sudah, jumlah = set(), 0
    for pdf_path in daftar:
        sidik = hashlib.md5(pdf_path.read_bytes()).hexdigest()
        if sidik in sudah:
            print(f"[DUPLIKAT] {pdf_path.name}: isi sama dengan file sebelumnya, dilewati")
            continue
        sudah.add(sidik)
        try:
            info, hasil = proses_pdf(pdf_path)
            if not hasil["transaksi"]:
                print(f"[KOSONG] {pdf_path.name}: tidak ada transaksi terbaca (PDF scan atau format berbeda)")
                continue
            bln, thn = bulan_tahun(info["Periode"])
            nama_file = f"REKOR_{info['No Rekening']}_{thn}-{bln:02d}.xlsx"
            n = tulis_excel(template, info, hasil, keluar / nama_file)
            ok = cek_laporan(hasil)
            print(f"[{'OK' if ok else 'CEK'}] {pdf_path.name} -> {nama_file} ({n} transaksi)")
            if not ok:
                print("      Total/jumlah transaksi tidak cocok dengan ringkasan di PDF, periksa hasilnya.")
            jumlah += 1
        except Exception as e:  # noqa: BLE001
            print(f"[ERROR] {pdf_path.name}: {e}")
    print(f"\nSelesai! {jumlah} file Excel disimpan di folder '{keluar}'.")


if __name__ == "__main__":
    main()
