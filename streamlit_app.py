import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import zipfile

st.set_page_config(page_title="XLS → XLSX Converter (Bulk)", layout="centered")

st.title("Konversi XLS ke XLSX (Bulk)")
st.caption("Konversi berbasis data (sheet & isi). Formatting Excel mungkin tidak ikut sepenuhnya.")

uploaded_files = st.file_uploader(
    "Upload file .xls (boleh banyak sekaligus)",
    type=["xls"],
    accept_multiple_files=True
)

def convert_xls_to_xlsx_bytes(xls_file) -> bytes:
    xls = pd.ExcelFile(xls_file, engine="xlrd")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name=sheet_name)
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    output.seek(0)
    return output.getvalue()

def make_zip(files_map: dict[str, bytes]) -> bytes:
    """files_map: {filename.xlsx: bytes}"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files_map.items():
            zf.writestr(name, data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

if not uploaded_files:
    st.info("Silakan upload file .xls untuk mulai konversi.")
    st.stop()

st.success(f"{len(uploaded_files)} file terdeteksi.")

# Tombol bulk convert
do_bulk = st.button("🚀 Konversi Semua & Download ZIP", type="primary")

# Bulk convert result storage
bulk_results = {}

if do_bulk:
    progress = st.progress(0)
    status = st.empty()

    for i, f in enumerate(uploaded_files, start=1):
        try:
            status.write(f"Memproses: **{f.name}** ({i}/{len(uploaded_files)})")
            xlsx_bytes = convert_xls_to_xlsx_bytes(f)
            out_name = f"{Path(f.name).stem}.xlsx"
            bulk_results[out_name] = xlsx_bytes
        except Exception as e:
            # tetap lanjut file berikutnya
            status.error(f"Gagal konversi {f.name}: {e}")

        progress.progress(int(i / len(uploaded_files) * 100))

    status.success("Selesai memproses bulk.")
    if bulk_results:
        zip_bytes = make_zip(bulk_results)
        st.download_button(
            label="⬇️ Download semua hasil (.zip)",
            data=zip_bytes,
            file_name="hasil_konversi_xls_ke_xlsx.zip",
            mime="application/zip"
        )
    else:
        st.warning("Tidak ada file yang berhasil dikonversi.")

st.divider()
st.subheader("Download per file (opsional)")

# Per-file convert + download (tanpa harus bulk)
for f in uploaded_files:
    col1, col2 = st.columns([3, 2], vertical_alignment="center")
    with col1:
        st.write(f"📄 **{f.name}**")
    with col2:
        try:
            xlsx_bytes = convert_xls_to_xlsx_bytes(f)
            out_name = f"{Path(f.name).stem}.xlsx"
            st.download_button(
                label="Download .xlsx",
                data=xlsx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{f.name}"
            )
        except Exception as e:
            st.error("Gagal")
            st.caption(str(e))
