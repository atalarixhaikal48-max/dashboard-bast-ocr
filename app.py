import base64
from datetime import datetime
import io
import os
import re

import fitz  # PyMuPDF
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from PIL import Image, ImageEnhance, ImageOps
import pytesseract
import streamlit as st
from streamlit_drawable_canvas import st_canvas

# =========================================================
# KONFIGURASI STREAMLIT
# =========================================================
st.set_page_config(page_title="Auto-Fill & Speedtest OCR", layout="wide")

# =========================================================
# KONFIGURASI PATH & INTEGRASI GOOGLE SHEETS
# =========================================================
SHEET_ID = "1nB917RJ0xxx"
CREDENTIALS_JSON = r"xxx-0f1692c4232b.json"

DEFAULT_TESSERACT_PATH = r"aaaa\Local\Programs\Tesseract-OCR\tesseract.exe"
if os.path.exists(DEFAULT_TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = DEFAULT_TESSERACT_PATH

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# =========================================================
# PEMETAAN KOLOM SESUAI HEADER GOOGLE SHEET
# =========================================================
FIELD_HEADER_MAP = {
    "wo_partner": "x",
    "woid": "x",
    "cust_name": "Customer Name",
    "phone": "Phone 1",
    "phone2": "Phone 2",
    "address": "Alamat",
    "city": "City",
    "zip": "ZIP",
    "email": "Email",
    "product": "Product",
    "tech_name": "NAMA_TEKNISI",
    "router": "Cable Suplier",
    "stb": "Route Type",
    "sn": "SN",
    "mac": "mMAC_ADDRESS",
    "loc": "DEVICE_LOCATION",
    "fat": "FAT_Code",
    "remark": "Remark",
    "cable_len": "Cable Length",
    "ds_24": "Speedtest 2,4 teknisi Download",
    "us_24": "Speedtest 2,4 teknisi Upload",
    "ds_5": "Speedtest 5,8 teknisi Download",
    "us_5": "Speedtest 5,8 teknisi Upload",
    "cust_ds_24": "Speedtest 2,4 Customer Download",
    "cust_us_24": "Speedtest 2,4 Customer Upload",
    "cust_ds_5": "Speedtest 5,8 Customer Download",
    "cust_us_5": "Speedtest 5,8 Customer Upload",
    "hub_pelanggan": "HUBUNGAN_PELANGGANTANDA_TANGAN_CUSTOMER",
    "sig_cust": "TANDA_TANGAN_CUSTOMER",
    "sig_tech": "TANDA_TANGAN_TEKNISI",
}

ALL_FIELDS = list(FIELD_HEADER_MAP.keys())
for field in ALL_FIELDS:
    if field not in st.session_state:
        st.session_state[field] = ""

if not st.session_state["tech_name"]:
    st.session_state["tech_name"] = "Dashboard User"

# =========================================================
# FUNGSI KONEKSI & READ DATA
# =========================================================


@st.cache_resource
def init_gspread_appsheet():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        CREDENTIALS_JSON, scope
    )
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).get_worksheet(0)


try:
    appsheet_sheet = init_gspread_appsheet()
except Exception:
    appsheet_sheet = None


def load_data_direct():
    if not appsheet_sheet:
        return pd.DataFrame()
    try:
        data = appsheet_sheet.get_all_records()
        df = pd.DataFrame(data)
        return df.fillna("")
    except Exception as e:
        st.error(f"Gagal mengambil data dari Google Sheet: {e}")
        return pd.DataFrame()


def get_col_value(row, target_header):
    if row is None:
        return ""
    target_clean = re.sub(r"[^a-zA-Z0-9]", "", target_header).lower()
    for col in row.index:
        col_clean = re.sub(r"[^a-zA-Z0-9]", "", str(col)).lower()
        if target_clean == col_clean or target_clean in col_clean:
            val = str(row[col]).strip()
            return "" if val.lower() == "nan" else val
    return ""


def apply_customer_data():
    selected_idx = st.session_state.get("selected_customer_idx")
    df = st.session_state.get("current_df")

    if selected_idx is not None and df is not None and selected_idx in df.index:
        row = df.loc[selected_idx]
        for field_key, header_name in FIELD_HEADER_MAP.items():
            val = get_col_value(row, header_name)
            if val:
                st.session_state[field_key] = val


def extract_number(text):
    match = re.search(r"[\d,.]+", text)
    if match:
        val = match.group(0).replace(",", ".")
        return val.strip(".")
    return "0"


def process_speedtest_ocr(pil_image):
    try:
        width, height = pil_image.size
        area_unduh = pil_image.crop(
            (width * 0.05, height * 0.10, width * 0.48, height * 0.25)
        )
        area_unggah = pil_image.crop(
            (width * 0.52, height * 0.10, width * 0.95, height * 0.25)
        )

        def enhance_image(image):
            image = image.convert("L")
            image = ImageOps.invert(image)
            enhancer = ImageEnhance.Contrast(image)
            return enhancer.enhance(2.0)

        txt_unduh = pytesseract.image_to_string(
            enhance_image(area_unduh), config="--psm 6 digits"
        )
        txt_unggah = pytesseract.image_to_string(
            enhance_image(area_unggah), config="--psm 6 digits"
        )

        return extract_number(txt_unduh), extract_number(txt_unggah)
    except Exception:
        return "0", "0"


def convert_canvas_to_base64(canvas_data):
    if canvas_data is not None and canvas_data.any():
        img = Image.fromarray(canvas_data.astype("uint8"))
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        return f"data:image/png;base64,{base64.b64encode(buffered.getvalue()).decode()}"
    return "Belum TTD"


# =========================================================
# AUTO DETECT & AUTO FILL ENGINE UNTUK PDF 
# =========================================================


def generate_baxx_pdf_custom(
    data,
    template_source,
    sig_cust_img=None,
    sig_tech_img=None,
):
    """Membuat PDF BAxx dengan koordinat yang diselaraskan ke template 612 x 1008 pt.

    Koordinat PyMuPDF dihitung dari sudut kiri-atas halaman. Jika template
    diganti dengan ukuran/versi berbeda, posisi perlu dikalibrasi ulang.
    """

    if isinstance(template_source, str):
        if not os.path.exists(template_source):
            raise FileNotFoundError(
                f"Template tidak ditemukan: {template_source}"
            )
        doc = fitz.open(template_source)
    else:
        template_source.seek(0)
        doc = fitz.open(
            stream=template_source.read(),
            filetype="pdf",
        )

    if len(doc) < 1:
        doc.close()
        raise ValueError("Template PDF tidak memiliki halaman.")

    page = doc[0]
    font_name = "helv"
    color = (0, 0, 0)
    layout_warnings = []

    # -----------------------------------------------------
    # FUNGSI PENULISAN TEKS DENGAN KOORDINAT TETAP
    # -----------------------------------------------------
    def write_auto_field(
        anchor_text,
        text_val,
        x_offset=120,
        width=200,
        font_size=8.0,
        align=fitz.TEXT_ALIGN_LEFT,
        fallback_rect=None,
    ):
        val_str = (
            str(text_val).strip()
            if text_val is not None
            and str(text_val).strip().lower()
            not in ("", "nan", "none")
            else "-"
        )

        rect = None

        # Koordinat tetap menjadi prioritas utama.
        if fallback_rect is not None:
            rect = fitz.Rect(*fallback_rect)

        # Pencarian label hanya menjadi alternatif.
        elif anchor_text:
            matches = page.search_for(anchor_text)

            if matches:
                ref = matches[0]
                rect = fitz.Rect(
                    ref.x0 + x_offset,
                    ref.y0 - 1,
                    ref.x0 + x_offset + width,
                    ref.y1 + 3,
                )

        if rect is None:
            layout_warnings.append(
                f"Posisi tidak ditemukan: "
                f"{anchor_text or val_str}"
            )
            return

        # Periksa batas halaman.
        page_rect = page.rect

        if (
            rect.x0 < page_rect.x0
            or rect.y0 < page_rect.y0
            or rect.x1 > page_rect.x1
            or rect.y1 > page_rect.y1
            or rect.width <= 0
            or rect.height <= 0
        ):
            layout_warnings.append(
                f"Koordinat di luar halaman: "
                f"{anchor_text or val_str}"
            )
            return

        # Cegah teks terlalu kecil atau tidak muat.
        current_size = font_size
        inserted = False

        while current_size >= 5.0:
            rc = page.insert_textbox(
                rect,
                val_str,
                fontsize=current_size,
                fontname=font_name,
                color=color,
                align=align,
                lineheight=1.1,
            )

            if rc >= 0:
                inserted = True
                break

            current_size -= 0.5

        if not inserted:
            layout_warnings.append(
                f"Teks tidak muat: {val_str!r}; "
                f"kotak {rect.width:.1f} x "
                f"{rect.height:.1f} pt"
            )

    # -----------------------------------------------------
    # 1. DETAIL WO
    # -----------------------------------------------------
    write_auto_field(
        "No WO:",
        data.get("woid"),
        x_offset=45,
        width=150,
        fallback_rect=(230, 108, 390, 122),
    )

    # -----------------------------------------------------
    # 2. DATA PELANGGAN
    # -----------------------------------------------------
    write_auto_field(
        "Nama Pelanggan",
        data.get("cust_name"),
        fallback_rect=(210, 151, 575, 166),
    )

    cust_id = data.get("wo_partner") or data.get("zip")

    write_auto_field(
        "No Cust ID",
        cust_id,
        fallback_rect=(210, 176, 575, 190),
    )

    full_address = str(data.get("address", "")).strip()
    city = str(data.get("city", "")).strip()

    if city and city.lower() not in full_address.lower():
        full_address = (
            f"{full_address}, {city}"
            if full_address
            else city
        )

    write_auto_field(
        "Alamat Rumah",
        full_address,
        font_size=7.0,
        fallback_rect=(210, 194, 575, 214),
    )

    write_auto_field(
        "Kode Pos",
        data.get("zip"),
        fallback_rect=(210, 218, 575, 234),
    )

    write_auto_field(
        "Nomer Telp 1",
        data.get("phone"),
        fallback_rect=(210, 243, 575, 260),
    )

    write_auto_field(
        "Nomer Telp 2",
        data.get("phone2"),
        fallback_rect=(210, 268, 575, 285),
    )

    write_auto_field(
        "Produk yang di subscribe",
        data.get("product"),
        fallback_rect=(210, 292, 575, 307),
    )

    write_auto_field(
        "E-mail",
        data.get("email"),
        font_size=7.5,
        fallback_rect=(210, 310, 575, 325),
    )

    write_auto_field(
        "Contact Person",
        data.get("cust_name"),
        fallback_rect=(210, 328, 575, 343),
    )

    # -----------------------------------------------------
    # 3. MATERIAL ROUTER
    # -----------------------------------------------------
    write_auto_field(
        "Tipe WIFI Router",
        data.get("router"),
        fallback_rect=(170, 390, 320, 407),
    )

    write_auto_field(
        "Serial Number",
        data.get("sn"),
        fallback_rect=(170, 414, 320, 430),
    )

    write_auto_field(
        "MAC Address",
        data.get("mac"),
        fallback_rect=(170, 436, 320, 452),
    )

    write_auto_field(
        "Supplier",
        data.get("router") or "XL Supplier",
        fallback_rect=(170, 458, 320, 474),
    )

    write_auto_field(
        "Device Location",
        data.get("loc"),
        fallback_rect=(170, 480, 320, 496),
    )

    write_auto_field(
        "FAT Code",
        data.get("fat"),
        fallback_rect=(170, 502, 320, 518),
    )

    # -----------------------------------------------------
    # 4. MATERIAL STB
    # -----------------------------------------------------
    write_auto_field(
        "Tipe STB",
        data.get("stb"),
        fallback_rect=(410, 390, 575, 407),
    )

    write_auto_field(
        None,
        data.get("sn"),
        fallback_rect=(410, 414, 575, 430),
    )

    write_auto_field(
        None,
        data.get("mac"),
        fallback_rect=(410, 436, 575, 452),
    )

    write_auto_field(
        None,
        data.get("stb") or "XL Supplier",
        fallback_rect=(410, 458, 575, 474),
    )

    write_auto_field(
        None,
        data.get("loc"),
        fallback_rect=(410, 480, 575, 496),
    )

    write_auto_field(
        None,
        data.get("fat"),
        fallback_rect=(410, 502, 575, 518),
    )

    # -----------------------------------------------------
    # 5. REMARK DAN PANJANG KABEL
    # -----------------------------------------------------
    write_auto_field(
        "Catatan Remark",
        data.get("remark") or "Installation Done",
        fallback_rect=(90, 620, 570, 634),
    )

    write_auto_field(
        "Cable Length Meter",
        data.get("cable_len"),
        fallback_rect=(95, 632, 180, 646),
    )

    # -----------------------------------------------------
    # 6. SPEEDTEST TEKNISI
    # -----------------------------------------------------
    speed_fields = [
        ("ds_24", (190, 657, 228, 670)),
        ("us_24", (190, 664, 228, 677)),
        ("ds_5", (190, 671, 228, 684)),
        ("us_5", (190, 679, 228, 692)),
    ]

    for key, rect in speed_fields:
        write_auto_field(
            None,
            data.get(key),
            align=fitz.TEXT_ALIGN_RIGHT,
            fallback_rect=rect,
        )

    # -----------------------------------------------------
    # 7. SPEEDTEST PELANGGAN
    # -----------------------------------------------------
    customer_speed_fields = [
        ("cust_ds_24", (190, 707, 228, 720)),
        ("cust_us_24", (190, 714, 228, 727)),
        ("cust_ds_5", (190, 721, 228, 734)),
        ("cust_us_5", (190, 728, 228, 741)),
    ]

    for key, rect in customer_speed_fields:
        write_auto_field(
            None,
            data.get(key),
            align=fitz.TEXT_ALIGN_RIGHT,
            fallback_rect=rect,
        )

    # -----------------------------------------------------
    # 8. TANGGAL BAxx
    # -----------------------------------------------------
    tgl_baxx = datetime.now().strftime("%d/%m/%Y")

    write_auto_field(
        "ditandatangani pada:",
        tgl_baxx,
        fallback_rect=(450, 740, 520, 754),
    )

    # -----------------------------------------------------
    # 9. TANDA TANGAN
    # -----------------------------------------------------
    def insert_signature(image, rect_values, label):
        if image is None:
            return

        try:
            image = image.convert("RGBA")

            # Buang latar putih dan potong area kosong.
            pixels = image.load()

            for y in range(image.height):
                for x in range(image.width):
                    r, g, b, a = pixels[x, y]

                    if min(r, g, b) >= 245:
                        pixels[x, y] = (255, 255, 255, 0)

            bbox = image.getbbox()

            if bbox:
                image = image.crop(bbox)

            image_buffer = io.BytesIO()
            image.save(image_buffer, format="PNG")

            page.insert_image(
                fitz.Rect(*rect_values),
                stream=image_buffer.getvalue(),
                keep_proportion=True,
                overlay=True,
            )

        except Exception as exc:
            layout_warnings.append(
                f"Tanda tangan {label} gagal: {exc}"
            )

    insert_signature(
        sig_cust_img,
        (100, 780, 230, 835),
        "pelanggan",
    )

    insert_signature(
        sig_tech_img,
        (385, 780, 515, 835),
        "teknisi",
    )

    # -----------------------------------------------------
    # 10. NAMA DI BAWAH TANDA TANGAN
    # -----------------------------------------------------
    def insert_centered_name(name, center_x, y):
        text = f"({str(name).strip() or '-'})"
        max_width = 160
        font_size = 8.0

        text_width = fitz.get_text_length(
            text,
            fontname=font_name,
            fontsize=font_size,
        )

        while text_width > max_width and font_size > 5.0:
            font_size -= 0.5
            text_width = fitz.get_text_length(
                text,
                fontname=font_name,
                fontsize=font_size,
            )

        page.insert_text(
            fitz.Point(center_x - text_width / 2, y),
            text,
            fontsize=font_size,
            fontname=font_name,
            color=color,
            overlay=True,
        )

    insert_centered_name(
        data.get("cust_name", ""),
        176,
        875,
    )

    insert_centered_name(
        data.get("tech_name", ""),
        463,
        875,
    )

    # -----------------------------------------------------
    # 11. PERINGATAN POSISI
    # -----------------------------------------------------
    if layout_warnings:
        st.warning(
            "Ada bagian PDF yang perlu diperiksa:"
        )

        for warning in layout_warnings:
            st.write(f"- {warning}")

    # -----------------------------------------------------
    # 12. SIMPAN KE BUFFER MEMORI
    # -----------------------------------------------------
    pdf_buffer = io.BytesIO()

    try:
        doc.save(pdf_buffer, garbage=4, deflate=True)
    finally:
        doc.close()

    pdf_buffer.seek(0)
    return pdf_buffer

# =========================================================
# DASHBOARD UTAMA STREAMLIT
# =========================================================
st.title("🚀 Auto-Fill BAxx PDF & OCR Speedtest Dashboard")

df_sheets = load_data_direct()
st.session_state["current_df"] = df_sheets

# --- Cari Data Pelanggan ---
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.subheader("🔎 Cari Data Pelanggan")
with col_head2:
    if st.button("🔄 Refresh Data"):
        st.rerun()

search_keyword = st.text_input("Ketik WOID / WO Partner / Customer Name:")

if search_keyword and not df_sheets.empty:
    mask = df_sheets.apply(
        lambda row: row.astype(str)
        .str.contains(search_keyword, case=False, na=False)
        .any(),
        axis=1,
    )
    filtered_df = df_sheets[mask]

    if not filtered_df.empty:
        st.success(f"Ditemukan {len(filtered_df)} data yang cocok.")

        options = {}
        for idx, row in filtered_df.iterrows():
            wo = get_col_value(row, "WOID") or "No WOID"
            partner = get_col_value(row, "WO Partner") or ""
            nama = get_col_value(row, "Customer Name") or "No Name"
            options[idx] = f"{wo} | {partner} | {nama}"

        st.selectbox(
            "Pilih Data Pelanggan:",
            options=list(options.keys()),
            format_func=lambda x: options[x],
            key="selected_customer_idx",
            on_change=apply_customer_data,
        )

        if (
            "last_search" not in st.session_state
            or st.session_state.last_search != search_keyword
        ):
            st.session_state.last_search = search_keyword
            apply_customer_data()
    else:
        st.warning("Data tidak ditemukan di Google Sheets.")

st.divider()

# --- Form Input Data ---
st.subheader("👤 Data Customer & Perangkat")
col1, col2 = st.columns(2)

with col1:
    st.text_input("WOID", key="woid")
    st.text_input("WO Partner", key="wo_partner")
    st.text_input("Customer Name", key="cust_name")
    st.text_input("Phone 1", key="phone")
    st.text_input("Phone 2", key="phone2")
    st.text_area("Alamat", key="address")
    st.text_input("City", key="city")
    st.text_input("ZIP", key="zip")

with col2:
    st.text_input("Email", key="email")
    st.text_input("Product", key="product")
    st.text_input("NAMA_TEKNISI", key="tech_name")
    st.text_input("Cable Suplier / Router", key="router")
    st.text_input("Route Type / STB", key="stb")
    st.text_input("SN (Serial Number)", key="sn")
    st.text_input("mMAC_ADDRESS", key="mac")
    st.text_input("DEVICE_LOCATION", key="loc")
    st.text_input("FAT_Code", key="fat")
    st.text_input("Remark", key="remark")
    st.text_input("Cable Length", key="cable_len")

st.divider()

# --- Speedtest & OCR ---
st.subheader("📸 Hasil Speedtest & Auto OCR")

uploaded_files = st.file_uploader(
    "Upload Gambar Speed Test (Foto 1 = 2.4 GHz, Foto 2 = 5.8 GHz)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True,
    key="bulk_uploader",
)

if st.button("🔍 PROSES OCR GAMBAR SPEEDTEST", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ Unggah minimal 1 gambar speedtest!")
    else:
        with st.spinner("Membaca data OCR dari gambar..."):
            if len(uploaded_files) >= 1:
                img1 = Image.open(uploaded_files[0])
                dl1, ul1 = process_speedtest_ocr(img1)
                st.session_state["ds_24"] = dl1
                st.session_state["us_24"] = ul1

            if len(uploaded_files) >= 2:
                img2 = Image.open(uploaded_files[1])
                dl2, ul2 = process_speedtest_ocr(img2)
                st.session_state["ds_5"] = dl2
                st.session_state["us_5"] = ul2

        st.success("✅ OCR Berhasil! Hasil telah terisi otomatis di bawah.")
        st.rerun()

st.markdown("**Perangkat Teknisi (Technician Device)**")
col_sp1, col_sp2 = st.columns(2)
with col_sp1:
    st.text_input("Speedtest 2,4 teknisi Download", key="ds_24")
    st.text_input("Speedtest 2,4 teknisi Upload", key="us_24")
with col_sp2:
    st.text_input("Speedtest 5,8 teknisi Download", key="ds_5")
    st.text_input("Speedtest 5,8 teknisi Upload", key="us_5")

st.markdown("**Perangkat Pelanggan (Customer Device)**")
col_csp1, col_csp2 = st.columns(2)
with col_csp1:
    st.text_input("Speedtest 2,4 Customer Download", key="cust_ds_24")
    st.text_input("Speedtest 2,4 Customer Upload", key="cust_us_24")
with col_csp2:
    st.text_input("Speedtest 5,8 Customer Download", key="cust_ds_5")
    st.text_input("Speedtest 5,8 Customer Upload", key="cust_us_5")

st.divider()

# --- Pengaturan Template PDF ---
st.subheader("📑 Template BAxx PDF")

pdf_files = [f for f in os.listdir(BASE_DIR) if f.endswith(".pdf")]
col_temp1, col_temp2 = st.columns(2)

with col_temp1:
    st.markdown("**1. Template Aktif Tersimpan**")
    if pdf_files:
        selected_pdf_file = st.selectbox(
            "Template yang sedang digunakan:", pdf_files, disabled=True
        )
        active_template_path = os.path.join(BASE_DIR, selected_pdf_file)

        if st.button("🗑️ Hapus Template Saat Ini"):
            os.remove(active_template_path)
            st.success(
                "Template berhasil dihapus. Silakan unggah template baru."
            )
            st.rerun()
    else:
        st.info("Belum ada template PDF tersimpan di server.")
        active_template_path = None

with col_temp2:
    st.markdown("**2. Unggah Template PDF Baru**")

    if pdf_files:
        st.warning(
            "🔒 **Template Terkunci!** Sudah ada template permanen yang tersimpan."
        )
        st.file_uploader(
            "Unggah File Template Baru (.pdf)",
            type=["pdf"],
            disabled=True,
            key="disabled_uploader",
        )
        uploaded_template = None
    else:
        uploaded_template = st.file_uploader(
            "Unggah File Template Baru (.pdf)",
            type=["pdf"],
            key="new_pdf_template_uploader",
        )

        if uploaded_template is not None:
            if st.button("💾 Simpan Permanen Template"):
                save_path = os.path.join(BASE_DIR, uploaded_template.name)
                with open(save_path, "wb") as f:
                    f.write(uploaded_template.getbuffer())
                st.success(
                    f"✅ Template tersimpan secara permanen: `{uploaded_template.name}`"
                )
                st.rerun()

st.divider()

# --- Tanda Tangan Digital & Tombol Action ---
st.subheader("✍️ Tanda Tangan Digital & Action")
sig_col1, sig_col2 = st.columns(2)

with sig_col1:
    st.markdown("**TANDA_TANGAN_CUSTOMER**")
    canvas_cust = st_canvas(
        fill_color="rgba(255,255,255,0)",
        stroke_width=2,
        stroke_color="#000",
        background_color="#FFF",
        height=130,
        width=280,
        drawing_mode="freedraw",
        key="canvas_customer",
    )

with sig_col2:
    st.markdown("**TANDA_TANGAN_TEKNISI**")
    canvas_tech = st_canvas(
        fill_color="rgba(255,255,255,0)",
        stroke_width=2,
        stroke_color="#000",
        background_color="#FFF",
        height=130,
        width=280,
        drawing_mode="freedraw",
        key="canvas_teknisi",
    )

btn_col1, btn_col2 = st.columns(2)

with btn_col1:
    if st.button("💾 SIMPAN / UPDATE DATA KE GOOGLE SHEET", type="primary"):
        current_woid = st.session_state.get("woid", "").strip()

        if not current_woid:
            st.warning(
                "⚠️ WOID tidak boleh kosong untuk menyimpan atau meng-update data!"
            )
        elif not appsheet_sheet:
            st.error("⚠️ Koneksi gspread ke Google Sheet gagal.")
        else:
            with st.spinner("Memproses update ke Google Sheet..."):
                try:
                    all_records = appsheet_sheet.get_all_values()
                    raw_headers = all_records[0]
                    headers_clean = [
                        re.sub(r"[^a-zA-Z0-9]", "", h).lower()
                        for h in raw_headers
                    ]

                    str_ttd_cust = convert_canvas_to_base64(
                        canvas_cust.image_data
                    )
                    str_ttd_tech = convert_canvas_to_base64(
                        canvas_tech.image_data
                    )

                    current_values = {
                        field_key: st.session_state.get(field_key, "")
                        for field_key in FIELD_HEADER_MAP
                    }
                    if "data:image" in str_ttd_cust:
                        current_values["sig_cust"] = str_ttd_cust
                    if "data:image" in str_ttd_tech:
                        current_values["sig_tech"] = str_ttd_tech

                    col_idx_map = {}
                    for field_key, header_target in FIELD_HEADER_MAP.items():
                        target_clean = re.sub(
                            r"[^a-zA-Z0-9]", "", header_target
                        ).lower()
                        for idx, h_clean in enumerate(headers_clean):
                            if (
                                target_clean == h_clean
                                or target_clean in h_clean
                            ):
                                col_idx_map[field_key] = idx
                                break

                    wo_col_idx = col_idx_map.get("woid", 0)
                    target_row_idx = None
                    for r_idx, row in enumerate(all_records[1:], start=2):
                        if (
                            len(row) > wo_col_idx
                            and row[wo_col_idx].strip().upper()
                            == current_woid.upper()
                        ):
                            target_row_idx = r_idx
                            break

                    if target_row_idx:
                        row_data = list(all_records[target_row_idx - 1])
                        while len(row_data) < len(headers_clean):
                            row_data.append("")
                        for field_key, val in current_values.items():
                            if field_key in col_idx_map and val != "":
                                row_data[col_idx_map[field_key]] = str(val)

                        cell_range = f"A{target_row_idx}:{gspread.utils.rowcol_to_a1(target_row_idx, len(row_data))}"
                        appsheet_sheet.update(cell_range, [row_data])
                        st.success(
                            f"🔄 **Berhasil Update!** Data WOID `{current_woid}` diperbarui."
                        )
                    else:
                        new_row = [""] * len(headers_clean)
                        for field_key, val in current_values.items():
                            if field_key in col_idx_map:
                                new_row[col_idx_map[field_key]] = str(val)

                        appsheet_sheet.append_row(
                            new_row, value_input_option="USER_ENTERED"
                        )
                        st.success(
                            f"✅ **Berhasil Simpan Baru!** Data WOID `{current_woid}` ditambahkan."
                        )

                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal memperbarui data Google Sheet: {e}")

with btn_col2:
    if st.button("📄 GENERATE PDF BAxx"):
        data_baxx = {
            k: st.session_state.get(k, "")
            for k in ALL_FIELDS
        }

        # Pastikan WOID tersedia.
        current_woid = str(
            data_baxx.get("woid", "")
        ).strip()

        if not current_woid:
            st.error("WOID wajib diisi.")
            st.stop()

        if uploaded_template is not None:
            template_to_use = uploaded_template

        elif (
            active_template_path
            and os.path.exists(active_template_path)
        ):
            template_to_use = active_template_path

        else:
            st.error("File template PDF tidak ditemukan.")
            st.stop()

        # Ambil tanda tangan dari canvas.
        img_cust = None
        img_tech = None

        if (
            canvas_cust.image_data is not None
            and canvas_cust.image_data.any()
        ):
            img_cust = Image.fromarray(
                canvas_cust.image_data.astype("uint8")
            )

        if (
            canvas_tech.image_data is not None
            and canvas_tech.image_data.any()
        ):
            img_tech = Image.fromarray(
                canvas_tech.image_data.astype("uint8")
            )

        try:
            pdf_file = generate_baxx_pdf_custom(
                data_baxx,
                template_source=template_to_use,
                sig_cust_img=img_cust,
                sig_tech_img=img_tech,
            )

            # Bersihkan nama file dari karakter ilegal.
            wo_filename = re.sub(
                r'[<>:"/\\|?*\x00-\x1f]',
                "_",
                current_woid,
            ).strip(" .")

            if not wo_filename:
                wo_filename = "BAxx"

            st.success(
                "PDF BAxx berhasil dibuat. "
                "Periksa hasilnya sebelum digunakan."
            )

            st.download_button(
                label="📥 DOWNLOAD PDF BAxx",
                data=pdf_file.getvalue(),
                file_name=f"BAxx_{wo_filename}.pdf",
                mime="application/pdf",
                key="download_baxx_pdf",
            )

        except Exception as e:
            st.error(f"Gagal membuat PDF BAxx: {e}")