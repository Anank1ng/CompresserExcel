import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import zipfile

st.set_page_config(page_title="XLS → XLSX Converter (Bulk)", layout="centered")

st.title("Konversi XLS ke XLSX (Bulk)")
st.caption("Auto-detect: XLS beneran / XLSX salah ekstensi / HTML export (disamarkan .xls).")

uploaded_files = st.file_uploader(
    "Upload file (boleh banyak sekaligus)",
    type=["xls", "xlsx", "htm", "html"],  # tetap bisa upload .xls palsu (HTML)
    accept_multiple_files=True
)

def _strip_bom_and_ws(b: bytes) -> bytes:
    b = b.lstrip()
    if b.startswith(b"\xef\xbb\xbf"):  # UTF-8 BOM
        b = b[3:].lstrip()
    return b

def sniff_file_type(data: bytes) -> str:
    head = data[:512]
    if head.startswith(b"PK\x03\x04"):
        return "xlsx"
    if head.startswith(b"\xD0\xCF\x11\xE0"):
        return "xls"
    low = _strip_bom_and_ws(head).lower()
    if low.startswith(b"<html") or low.startswith(b"<!doctype") or low.startswith(b"<?xml") or low.startswith(b"<table"):
        return "html"
    if b"<html" in low:
        return "html"
    return "unknown"

def convert_to_xlsx_bytes(uploaded_file) -> tuple[bytes, str]:
    data = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    ftype = sniff_file_type(data)

    out = BytesIO()

    # 1) XLSX (atau file XLSX yang salah ekstensi)
    if ftype == "xlsx":
        xls = pd.ExcelFile(BytesIO(data), engine="openpyxl")
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            for sh in xls.sheet_names:
                xls.parse(sh).to_excel(w, sheet_name=sh[:31], index=False)
        out.seek(0)
        return out.getvalue(), "xlsx (detected)"

    # 2) HTML export yang disamarkan jadi .xls
    if ftype == "html":
        html_text = data.decode("utf-8", errors="ignore")
        tables = pd.read_html(html_text)  # butuh lxml/bs4
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            for i, df in enumerate(tables, start=1):
                df.to_excel(w, sheet_name=f"Table{i}"[:31], index=False)
        out.seek(0)
        return out.getvalue(), "html export (detected)"

    # 3) XLS beneran (BIFF)
    if ftype == "xls":
        xls = pd.ExcelFile(BytesIO(data), engine="xlrd")
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            for sh in xls.sheet_names:
                xls.parse(sh).to_excel(w, sheet_name=sh[:31], index=False)
        out.seek(0)
        return out.getvalue(), "xls (detected)"

    # 4) Unknown: coba XLS dulu, kalau gagal coba HTML
    try:
        xls = pd.ExcelFile(BytesIO(data), engine="xlrd")
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            for sh in xls.sheet_names:
                xls.parse(sh).to_excel(w, sheet_name=sh[:31], index=False)
        out.seek(0)
        return out.getvalue(), "xls (fallback)"
    except Exception:
        html_text = data.decode("utf-8", errors="ignore")
        tables = pd.read_html(html_text)
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            for i, df in enumerate(tables, start=1):
                df.to_excel(w, sheet_name=f"Table{i}"[:31], index=False)
        out.seek(0)
        return out.getvalue(), "html (fallback)"

def make_zip(files_map: dict[str, bytes]) -> bytes:
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files_map.items():
            zf.writestr(name, data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

if not uploaded_files:
    st.info("Silakan upload file untuk mulai konversi.")
    st.stop()

st.success(f"{len(uploaded_files)} file terdeteksi.")
do_bulk = st.button("🚀 Konversi Semua & Download ZIP", type="primary")

if do_bulk:
    progress = st.progress(0)
    status = st.empty()
    bulk_results: dict[str, bytes] = {}

    for i, f in enumerate(uploaded_files, start=1):
        try:
            status.write(f"Memproses: **{f.name}** ({i}/{len(uploaded_files)})")
            xlsx_bytes, detected = convert_to_xlsx_bytes(f)
            out_name = f"{Path(f.name).stem}.xlsx"
            bulk_results[out_name] = xlsx_bytes
            st.caption(f"✅ {f.name} → {out_name} ({detected})")
        except Exception as e:
            st.error(f"❌ Gagal konversi {f.name}: {e}")

        progress.progress(int(i / len(uploaded_files) * 100))

    if bulk_results:
        zip_bytes = make_zip(bulk_results)
        st.download_button(
            label="⬇️ Download semua hasil (.zip)",
            data=zip_bytes,
            file_name="hasil_konversi_xlsx.zip",
            mime="application/zip"
        )
    else:
        st.warning("Tidak ada file yang berhasil dikonversi.")

st.divider()
st.subheader("Download per file (opsional)")

for f in uploaded_files:
    col1, col2 = st.columns([3, 2], vertical_alignment="center")
    with col1:
        st.write(f"📄 **{f.name}**")
    with col2:
        try:
            xlsx_bytes, detected = convert_to_xlsx_bytes(f)
            out_name = f"{Path(f.name).stem}.xlsx"
            st.download_button(
                label=f"Download .xlsx",
                data=xlsx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{f.name}"
            )
            st.caption(detected)
        except Exception as e:
            st.error("Gagal")
            st.caption(str(e))
