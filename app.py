"""
Aplikasi Streamlit: PDF mutasi rekening giro BCA -> Excel (template REKOR).
1 file PDF = 1 file Excel.

Jalankan lokal :  streamlit run app.py
File yang harus satu folder:
    app.py, pdf_ke_rekor_bca.py, REKOR_2026.xlsx, requirements.txt
"""

import hashlib
import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from pdf_ke_rekor_bca import (
    bulan_tahun,
    cek_laporan,
    keterangan_bank,
    keterangan_lengkap,
    nama_lawan,
    proses_pdf,
    tulis_excel,
    validasi,
)

TEMPLATE_BAWAAN = Path(__file__).resolve().parent / "REKOR_2026.xlsx"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def rp(x):
    """Format angka gaya Indonesia: 1.234.567,89"""
    if x is None:
        return "-"
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner=False, ttl=3600, max_entries=30)
def olah_pdf(data: bytes, template_bytes: bytes):
    """Ekstrak satu PDF dan buat file Excel-nya (di memori, tidak disimpan ke disk)."""
    info, hasil = proses_pdf(io.BytesIO(data))
    if not hasil["transaksi"]:
        raise ValueError("Tidak ada transaksi yang terbaca. PDF mungkin hasil scan atau "
                         "bukan format mutasi giro BCA.")

    bln, thn = bulan_tahun(info["Periode"])
    nama_file = f"REKOR_{info['No Rekening']}_{thn}-{bln:02d}.xlsx"

    buf = io.BytesIO()
    tulis_excel(io.BytesIO(template_bytes), info, hasil, buf)

    trx = hasil["transaksi"]
    df = pd.DataFrame({
        "No": range(1, len(trx) + 1),
        "Tanggal": [t["Tanggal"].date() for t in trx],
        "Keterangan Bank": [keterangan_bank(t["Jenis Transaksi"], t["jenis_extra"]) for t in trx],
        "Keterangan 1 (Lengkap)": [keterangan_lengkap(t["detail"]) for t in trx],
        "Nama Lawan Transaksi": [nama_lawan(t["Jenis Transaksi"], t["detail"]) for t in trx],
        "Debit (Masuk)": [t["Kredit"] for t in trx],
        "Kredit (Keluar)": [t["Debit"] for t in trx],
        "Saldo Bank": [t["Saldo"] for t in trx],
    })
    v = validasi(hasil)
    return {
        "info": info,
        "nama_file": nama_file,
        "excel": buf.getvalue(),
        "ok": cek_laporan(hasil),
        "saldo_awal": hasil["saldo_awal"],
        "saldo_akhir": v["saldo_akhir_pdf"],
        "total_masuk": df["Debit (Masuk)"].sum(),
        "total_keluar": df["Kredit (Keluar)"].sum(),
        "df": df,
    }


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


def main():
    st.set_page_config(page_title="PDF Mutasi BCA ke Excel", page_icon="📄", layout="wide")
    st.title("📄 PDF Mutasi Giro BCA → Excel REKOR")
    st.caption("Unggah PDF mutasi rekening giro BCA. Setiap PDF menghasilkan satu file Excel "
               "sesuai template.")

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
        st.markdown(
            "**Catatan**\n\n"
            "- Uang masuk (CR) → kolom **Debit**; uang keluar (DB) → kolom **Kredit**.\n"
            "- Keterangan 1 diisi keterangan transaksi lengkap; Jenis Transaksi dan No Faktur Pajak dikosongkan.\n"
            "- Hanya untuk PDF digital format giro BCA (bukan hasil scan).\n"
            "- File diproses di memori dan tidak disimpan oleh aplikasi."
        )

    # ---------------- upload
    files = st.file_uploader("1. Unggah file PDF mutasi (boleh lebih dari satu)",
                             type=["pdf"], accept_multiple_files=True)
    if not files:
        st.info("Belum ada file. Unggah satu atau beberapa PDF untuk memulai.")
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

        try:
            r = olah_pdf(data, template_bytes)
        except Exception as e:  # noqa: BLE001
            st.error(f"**{f.name}**: gagal diproses. {e}")
            continue

        r["nama_file"] = nama_unik(r["nama_file"], terpakai)
        hasil_ok.append(r)
        info = r["info"]
        ikon = "✅" if r["ok"] else "⚠️"
        with st.expander(f"{ikon} {f.name}  →  {r['nama_file']}", expanded=len(files) == 1 or not r["ok"]):
            if r["ok"]:
                st.success("Total dan jumlah transaksi cocok dengan ringkasan di PDF.")
            else:
                st.warning("Total/jumlah transaksi TIDAK cocok dengan ringkasan di PDF. "
                           "Periksa hasilnya sebelum dipakai.")
            c = st.columns(4)
            c[0].metric("Periode", info["Periode"] or "-")
            c[1].metric("No. Rekening", info["No Rekening"] or "-")
            c[2].metric("Nama", info["Nama"] or "-")
            c[3].metric("Jumlah Transaksi", len(r["df"]))
            c = st.columns(4)
            c[0].metric("Saldo Awal", rp(r["saldo_awal"]))
            c[1].metric("Total Masuk (Debit)", rp(r["total_masuk"]))
            c[2].metric("Total Keluar (Kredit)", rp(r["total_keluar"]))
            c[3].metric("Saldo Akhir (PDF)", rp(r["saldo_akhir"]))

            st.dataframe(
                r["df"], use_container_width=True, hide_index=True, height=350,
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
