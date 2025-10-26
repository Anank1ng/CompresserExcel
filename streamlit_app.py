# streamlit_app.py
import io
import os
import zipfile
from datetime import datetime
import importlib

import streamlit as st


st.set_page_config(page_title="Excel Shrinker", page_icon="🧹", layout="centered")
st.title("🧹 Excel Shrinker — Kecilkan ukuran tanpa mengubah isi")
st.caption("Pilih metode kompresi sesuai kebutuhanmu. Mode Lossless aman untuk semua file .xlsx/.xlsm.")


# ===== Utilities =====
def human_size(num_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < 1024.0:
            return f"{num_bytes:,.2f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:,.2f} TB"


def detect_vba(xlsx_bytes: bytes) -> bool:
    try:
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as z:
            return any(name.lower().startswith("xl/") and name.lower().endswith("vbaProject.bin".lower())
                       for name in z.namelist())
    except zipfile.BadZipFile:
        return False


def recompress_xlsx(xbytes: bytes, compresslevel: int = 9) -> bytes:
    """Lossless: Re-zip semua entry. Tidak mengubah isi/struktur workbook."""
    src_io = io.BytesIO(xbytes)
    out_io = io.BytesIO()
    with zipfile.ZipFile(src_io, "r") as zin, zipfile.ZipFile(out_io, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            is_binary = info.filename.lower().endswith((
                ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emf", ".wmf", ".bin"
            ))
            if is_binary:
                zout.writestr(info.filename, data, compress_type=zipfile.ZIP_STORED)
            else:
                try:
                    zout.writestr(
                        info.filename,
                        data,
                        compress_type=zipfile.ZIP_DEFLATED,
                        compresslevel=int(compresslevel),
                    )
                except TypeError:
                    # Python tua tanpa argumen compresslevel
                    zout.writestr(info.filename, data, compress_type=zipfile.ZIP_DEFLATED)
    out_io.seek(0)
    return out_io.getvalue()


def have_openpyxl() -> bool:
    try:
        importlib.import_module("openpyxl")
        return True
    except ImportError:
        return False


def minify_xlsx(xbytes: bytes, keep_number_formats: bool = True) -> bytes:
    """
    Copy nilai & formula saja ke workbook baru (write_only) -> buang styling, objek, gambar, dsb.
    Opsi: jaga number format agar tanggal/angka tetap terbaca. (Tidak mendukung file ber-makro)
    """
    try:
        from openpyxl import load_workbook, Workbook
        from openpyxl.cell.cell import WriteOnlyCell
    except ImportError as ie:
        raise RuntimeError(
            "Mode Minify memerlukan paket 'openpyxl'. Instal dengan:\n"
            "  py -m pip install openpyxl\n"
            "atau\n"
            "  python -m pip install openpyxl"
        ) from ie

    src = io.BytesIO(xbytes)
    wb_src = load_workbook(src, data_only=False, keep_links=True)
    wb_dst = Workbook(write_only=True)

    first = True
    for ws in wb_src.worksheets:
        if first:
            ws_new = wb_dst.active
            ws_new.title = ws.title
            first = False
        else:
            ws_new = wb_dst.create_sheet(title=ws.title)

        try:
            ws_new.sheet_state = ws.sheet_state
        except Exception:
            pass

        max_row = ws.max_row or 1
        max_col = ws.max_column or 1

        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            out_row = []
            for c in row:
                val = c.value
                wcell = WriteOnlyCell(ws_new, value=val)

                if getattr(c, "hyperlink", None):
                    link = getattr(c.hyperlink, "target", None) or str(c.hyperlink)
                    if link:
                        wcell.hyperlink = link

                if keep_number_formats:
                    try:
                        wcell.number_format = c.number_format
                    except Exception:
                        pass

                out_row.append(wcell)
            ws_new.append(out_row)

    out = io.BytesIO()
    wb_dst.save(out)
    out.seek(0)
    return out.getvalue()


# ===== UI =====
uploaded = st.file_uploader(
    "Unggah file Excel (.xlsx atau .xlsm)", type=["xlsx", "xlsm"], accept_multiple_files=False
)

mode = st.radio(
    "Pilih metode kompresi:",
    options=["Lossless (Recompress ZIP)", "Minify (Value & Formula Only)"],
    help=(
        "• Lossless: Tidak mengubah isi/format/objek sama sekali; hanya kompres ulang ZIP di dalam .xlsx/.xlsm.\n"
        "• Minify: Salin nilai & formula ke workbook baru (tanpa styling/objek) agar lebih kecil."
    ),
)

keep_formats = st.checkbox(
    "Pertahankan format angka/tanggal (disarankan untuk Minify)",
    value=True,
)

compress_lvl = st.slider(
    "Tingkat kompresi (Lossless)", min_value=1, max_value=9, value=9,
    help="Hanya berlaku untuk mode Lossless."
)

if uploaded is not None:
    name, ext = os.path.splitext(uploaded.name)
    ext = ext.lower()
    in_bytes = uploaded.read()
    in_size = len(in_bytes)
    st.write(f"**Ukuran asli:** {human_size(in_size)}")

    has_macro = detect_vba(in_bytes) if ext == ".xlsm" else False

    # Info dini kalau user pilih Minify tetapi openpyxl belum terpasang
    if mode.startswith("Minify") and not have_openpyxl():
        st.warning(
            "Mode **Minify** memerlukan paket `openpyxl`.\n\n"
            "Install di Terminal/Command Prompt:\n"
            "`py -m pip install openpyxl` atau `python -m pip install openpyxl`."
        )

    if mode.startswith("Minify") and
