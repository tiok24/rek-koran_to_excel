"""
Aplikasi Streamlit: mutasi rekening (PDF BCA / CIMB Niaga / Danamon / BRI, atau CSV) -> Excel
template REKOR.  1 file = 1 file Excel.

Jalankan lokal : streamlit run app.py

File yang harus satu folder / struktur:
    app.py, pdf_ke_rekor_bca.py, REKOR_2026.xlsx, requirements.txt, bank_parsers/
"""

import hashlib
import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from bank_parsers import BANK_PDF, csv_parser
from bank_parsers.pipeline import csv_bca_terdeteksi, olah_csv, olah_csv_bca, olah_pdf

TEMPLATE_BAWAAN = Path(__file__).resolve().parent / "REKOR_2026.xlsx"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
OTOMATIS = "Otomatis (deteksi)"
KOSONG = "— tidak ada —"
LABEL_KOLOM = {
    "tanggal": "Tanggal", "keterangan": "Keterangan / uraian",
    "debit": "Debit (uang keluar)", "kredit": "Kredit (uang masuk)",
    "saldo": "Saldo", "jumlah": "Jumlah (1 kolom, bertanda +/−)",
    "tipe": "Tipe DB/CR (untuk kolom jumlah)",
}


def tampil_df(df, **kw):
    """st.dataframe selebar kontainer; kompatibel dengan Streamlit lama maupun baru."""
    try:
        st.dataframe(df, width="stretch", **kw)
    except Exception:  # noqa: BLE001  (Streamlit < 1.49 tidak mengenal width="stretch")
        st.dataframe(df, use_container_width=True, **kw)


def rp(x):
    """Format angka gaya Indonesia: 1.234.567,89"""
    if x is None:
        return "-"
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner=False, ttl=3600, max_entries=30)
def olah_pdf_cache(data: bytes, template_bytes: bytes, bank):
    """Ekstrak satu PDF dan buat file Excel-nya (di memori, tidak disimpan ke disk)."""
    return olah_pdf(data, template_bytes, bank)


@st.cache_data(show_spinner=False, ttl=3600, max_entries=30)
def olah_csv_cache(data: bytes, template_bytes: bytes, pemetaan: dict, opsi: dict):
    return olah_csv(data, template_bytes, pemetaan, opsi)


@st.cache_data(show_spinner=False, ttl=3600, max_entries=30)
def olah_csv_bca_cache(data: bytes, template_bytes: bytes):
    return olah_csv_bca(data, template_bytes)


def nama_unik(nama, terpakai):
    if nama not in terpakai:
        terpakai.add(nama)
        return nama
    dasar, ekstensi = nama.rsplit(".", 1)
    i = 2
    while f"{dasar}_{i}.{ekstensi}" in terpakai:
        i += 1
    baru = f"{dasar}_{i}.{ekstensi}"
    terpakai.add(baru)
    return baru


def panel_csv(f, data, sidik):
    """Pengaturan konversi CSV untuk satu file. Mengembalikan (pemetaan, opsi) atau None."""
    try:
        header, baris, pemisah = csv_parser.muat_csv(data)
    except Exception as e:  # noqa: BLE001
        st.error(f"**{f.name}**: CSV tidak bisa dibaca. {e}")
        return None

    st.caption(f"Pemisah terdeteksi: `{pemisah!r}` · {len(baris)} baris data · "
               f"{len(header)} kolom. Periksa pemetaan kolom di bawah.")
    tampil_df(pd.DataFrame(baris[:8], columns=header), hide_index=True)

    tebakan = csv_parser.tebak_pemetaan(header)
    pilihan = [KOSONG] + header
    k = lambda nama: f"csv_{sidik}_{nama}"          # noqa: E731

    pemetaan = {}
    kolom = st.columns(4)
    for n, kunci in enumerate(csv_parser.KUNCI):
        awal = pilihan.index(tebakan[kunci]) if tebakan[kunci] in header else 0
        dipilih = kolom[n % 4].selectbox(LABEL_KOLOM[kunci], pilihan, index=awal, key=k(kunci))
        pemetaan[kunci] = None if dipilih == KOSONG else dipilih

    c = st.columns(4)
    nama = c[0].text_input("Nama pemilik rekening", key=k("nama"))
    no_rek = c[1].text_input("No. rekening", key=k("norek"))
    label = c[2].text_input("Label bank", value="BANK", key=k("bank"))
    fmt = c[3].selectbox("Format angka", ["Otomatis", "1.234,56 (Indonesia)", "1,234.56 (Inggris)"],
                         key=k("fmt"))
    c = st.columns(3)
    tukar = c[0].checkbox("Tukar arah debit ↔ kredit", key=k("tukar"),
                          help="Centang bila di CSV Anda kolom 'debit' berarti uang MASUK.")
    manual = c[1].checkbox("Isi saldo awal manual", key=k("manual"),
                           help="Bila kosong, saldo awal dihitung dari saldo baris pertama.")
    saldo_awal = c[2].number_input("Saldo awal", value=0.0, format="%.2f", key=k("sa"),
                                   disabled=not manual)

    opsi = {
        "nama": nama.strip(), "no_rekening": no_rek.strip(), "label_bank": label.strip(),
        "format_angka": {"Otomatis": "auto", "1.234,56 (Indonesia)": "id",
                         "1,234.56 (Inggris)": "en"}[fmt],
        "tukar_arah": tukar, "saldo_awal": float(saldo_awal) if manual else None,
        "pemisah": pemisah,
    }
    return pemetaan, opsi


def main():
    st.set_page_config(page_title="Mutasi Bank ke Excel REKOR", page_icon="📄", layout="wide")
    st.title("📄 Mutasi Rekening → Excel REKOR")
    st.caption("Unggah PDF mutasi **BCA, CIMB Niaga, Danamon, BRI** atau file **CSV** "
               "(CSV KlikBCA Bisnis dikenali otomatis). "
               "Setiap file menghasilkan satu file Excel sesuai template.")

    # ---------------- sidebar
    with st.sidebar:
        st.header("Template Excel")
        template_upload = st.file_uploader("Pakai template lain (opsional)", type=["xlsx"])
        if template_upload is not None:
            template_bytes = template_upload.getvalue()
            st.success(f"Memakai template: {template_upload.name}")
        elif TEMPLATE_BAWAAN.exists():
            template_bytes = TEMPLATE_BAWAAN.read_bytes()
            st.info(f"Memakai template bawaan: {TEMPLATE_BAWAAN.name}")
        else:
            st.error("Template bawaan (REKOR_2026.xlsx) tidak ditemukan. "
                     "Unggah template di atas.")
            st.stop()

        st.divider()
        st.header("Bank untuk PDF")
        bank_pilihan = st.selectbox("Bank", [OTOMATIS] + BANK_PDF,
                                    help="Biarkan otomatis; pilih manual hanya bila bank salah terdeteksi.")
        bank_pilihan = None if bank_pilihan == OTOMATIS else bank_pilihan

        st.divider()
        st.markdown(
            "**Catatan**\n\n"
            "- Uang masuk (CR) → kolom **Debit**; uang keluar (DB) → kolom **Kredit**.\n"
            "- Keterangan Bank = jenis/uraian transaksi; Keterangan 1 = detail lengkap; "
            "Keterangan 2 = nama lawan transaksi (bila bisa dikenali).\n"
            "- Hanya untuk PDF digital (bukan hasil scan).\n"
            "- File diproses di memori dan tidak disimpan oleh aplikasi."
        )

    # ---------------- upload
    files = st.file_uploader("1. Unggah file mutasi: PDF atau CSV (boleh lebih dari satu)",
                             type=["pdf", "csv", "txt"], accept_multiple_files=True)
    if not files:
        st.info("Belum ada file. Unggah satu atau beberapa PDF/CSV untuk memulai.")
        return

    st.subheader("2. Hasil ekstraksi")
    hasil_ok, terpakai, sudah = [], set(), set()
    progres = st.progress(0.0, text="Memproses...")

    for i, f in enumerate(files, start=1):
        data = f.getvalue()
        progres.progress(i / len(files), text=f"Memproses {f.name} ({i}/{len(files)})")
        sidik = hashlib.md5(data).hexdigest()
        if sidik in sudah:
            st.warning(f"**{f.name}**: isinya sama dengan file yang sudah diunggah, dilewati.")
            continue
        sudah.add(sidik)

        la_csv = f.name.lower().endswith((".csv", ".txt"))
        try:
            if la_csv and csv_bca_terdeteksi(data):
                # CSV KlikBCA Bisnis: nama, no. rekening, periode & saldo dibaca dari file,
                # tidak perlu pemetaan kolom manual.
                r = olah_csv_bca_cache(data, template_bytes)
            elif la_csv:
                with st.expander(f"⚙️ Pengaturan CSV — {f.name}", expanded=True):
                    pengaturan = panel_csv(f, data, sidik)
                if pengaturan is None:
                    continue
                r = olah_csv_cache(data, template_bytes, *pengaturan)
            else:
                r = olah_pdf_cache(data, template_bytes, bank_pilihan)
        except Exception as e:  # noqa: BLE001
            st.error(f"**{f.name}**: gagal diproses. {e}")
            continue

        r["nama_file"] = nama_unik(r["nama_file"], terpakai)
        hasil_ok.append(r)

        info = r["info"]
        if r["ok"]:
            ikon = "✅"
        elif not r["bisa_diverifikasi"]:
            ikon = "ℹ️"
        else:
            ikon = "⚠️"
        with st.expander(f"{ikon} [{r['bank']}] {f.name} → {r['nama_file']}",
                         expanded=len(files) == 1 or ikon == "⚠️"):
            if r["ok"]:
                st.success("Total, jumlah transaksi dan saldo berjalan cocok dengan ringkasan/saldo di dokumen.")
            elif not r["bisa_diverifikasi"]:
                st.info("Dokumen tidak memuat kolom saldo/ringkasan, jadi hasil tidak bisa diverifikasi "
                        "otomatis. Periksa tabel di bawah.")
            else:
                st.warning("Hasil TIDAK cocok dengan dokumen asli. Periksa sebelum dipakai:\n\n"
                           + "\n".join(f"- {p}" for p in r["pesan"]))
            if r["urutan_dibalik"]:
                st.caption("CSV tercatat terbaru-di-atas; urutan sudah dibalik menjadi kronologis.")
            for c in r.get("catatan", []):
                st.caption(f"📝 {c}")
            if r["dilewati"]:
                st.caption(f"{r['dilewati']} baris tanpa tanggal/nominal valid dilewati "
                           "(judul, total, catatan kaki, dll).")

            c = st.columns(4)
            c[0].metric("Periode", info["Periode"] or "-")
            c[1].metric("No. Rekening", info["No Rekening"] or "-")
            c[2].metric("Nama", info["Nama"] or "-")
            c[3].metric("Jumlah Transaksi", len(r["df"]))

            c = st.columns(4)
            c[0].metric("Saldo Awal", rp(r["saldo_awal"]))
            c[1].metric("Total Masuk (Debit)", rp(r["total_masuk"]))
            c[2].metric("Total Keluar (Kredit)", rp(r["total_keluar"]))
            c[3].metric("Saldo Akhir", rp(r["saldo_akhir"]))

            tampil_df(
                r["df"], hide_index=True, height=350,
                column_config={
                    "Tanggal": st.column_config.DateColumn(format="DD/MM/YYYY"),
                    "Debit (Masuk)": st.column_config.NumberColumn(format="%.2f"),
                    "Kredit (Keluar)": st.column_config.NumberColumn(format="%.2f"),
                    "Saldo Bank": st.column_config.NumberColumn(format="%.2f"),
                },
            )
            st.download_button("⬇️ Unduh Excel", data=r["excel"], file_name=r["nama_file"],
                               mime=MIME_XLSX, key=f"unduh_{i}")
    progres.empty()

    # ---------------- unduh semua
    if hasil_ok:
        st.subheader("3. Unduh")
        st.write(f"{len(hasil_ok)} file Excel siap diunduh.")
        if len(hasil_ok) > 1:
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                for r in hasil_ok:
                    z.writestr(r["nama_file"], r["excel"])
            st.download_button("⬇️ Unduh semua (ZIP)", data=zbuf.getvalue(),
                               file_name="hasil_rekor.zip", mime="application/zip",
                               key="unduh_zip")


if __name__ == "__main__":
    main()
