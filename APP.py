import streamlit as st
import pandas as pd
import re
import tempfile
import unicodedata
import urllib.request
import difflib
import smtplib
import ssl
import dns.resolver
import time
import zipfile
from email.message import EmailMessage
from datetime import date
from pathlib import Path
from io import BytesIO
from fpdf import FPDF
from fpdf.enums import XPos, YPos

# =====================================================
# 1. CẤU HÌNH TRANG
# =====================================================
st.set_page_config(
    page_title="Phần mềm tính tiền lương",
    page_icon="💰",
    layout="wide"
)

# =====================================================
# 2. LOGO - KHÔNG LÀM APP BỊ LỖI NẾU THIẾU ẢNH
# =====================================================
THU_MUC_GOC = Path(__file__).resolve().parent
DUONG_DAN_LOGO = THU_MUC_GOC / "LOGO.jpg"

if DUONG_DAN_LOGO.exists():
    st.image(str(DUONG_DAN_LOGO), width=200)

# =====================================================
# 3. HÀM ĐỊNH DẠNG
# =====================================================
def dinh_dang_tien(so_tien):
    return f"{so_tien:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")


MAU_EMAIL = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")


def email_hop_le(email):
    return bool(MAU_EMAIL.match(str(email).strip()))


def bo_dau(chuoi):
    # Chuyển tiếng Việt có dấu thành không dấu (dùng cho tên file / font dự phòng)
    chuoi = str(chuoi).replace("đ", "d").replace("Đ", "D")
    chuoi = unicodedata.normalize("NFD", chuoi)
    return "".join(c for c in chuoi if unicodedata.category(c) != "Mn")


def ten_file_an_toan(chuoi):
    chuoi = re.sub(r"[^A-Za-z0-9]+", "_", bo_dau(chuoi)).strip("_")
    return chuoi or "nhan_vien"

# =====================================================
# 4. HÀM TÍNH THUẾ TNCN
# Giữ nguyên biểu thuế theo code ban đầu của người dùng
# =====================================================
def tinh_thue_tncn(thu_nhap_tinh_thue):
    if thu_nhap_tinh_thue <= 0:
        return 0.0
    elif thu_nhap_tinh_thue <= 10:
        return thu_nhap_tinh_thue * 0.05
    elif thu_nhap_tinh_thue <= 30:
        return thu_nhap_tinh_thue * 0.10 - 0.5
    elif thu_nhap_tinh_thue <= 60:
        return thu_nhap_tinh_thue * 0.20 - 3.5
    elif thu_nhap_tinh_thue <= 100:
        return thu_nhap_tinh_thue * 0.30 - 9.5
    else:
        return thu_nhap_tinh_thue * 0.35 - 14.5

# =====================================================
# 5. HÀM TÍNH LƯƠNG CHO MỘT NHÂN VIÊN
# =====================================================
def tinh_luong(nhan_vien):
    luong_co_ban = float(nhan_vien["Lương cơ bản"])
    thuong = float(nhan_vien["Tiền thưởng"])
    phu_cap_co_bhxh = float(nhan_vien["Phụ cấp tính BHXH"])
    phu_cap_khong_bhxh = float(nhan_vien["Phụ cấp không tính BHXH"])
    so_nguoi_phu_thuoc = int(nhan_vien["Số người phụ thuộc"])

    tong_thu_nhap = (
        luong_co_ban
        + thuong
        + phu_cap_co_bhxh
        + phu_cap_khong_bhxh
    )

    tien_bao_hiem = (luong_co_ban + phu_cap_co_bhxh) * 0.105

    giam_tru_ban_than = 15.5
    giam_tru_nguoi_phu_thuoc = 6.2 * so_nguoi_phu_thuoc

    thu_nhap_tinh_thue = (
        tong_thu_nhap
        - giam_tru_ban_than
        - giam_tru_nguoi_phu_thuoc
        - tien_bao_hiem
    )

    thue_tncn = tinh_thue_tncn(thu_nhap_tinh_thue)
    thuc_lanh = tong_thu_nhap - tien_bao_hiem - thue_tncn

    return {
        "Tổng thu nhập": tong_thu_nhap,
        "Bảo hiểm người lao động đóng": tien_bao_hiem,
        "Giảm trừ bản thân": giam_tru_ban_than,
        "Giảm trừ người phụ thuộc": giam_tru_nguoi_phu_thuoc,
        "Thu nhập tính thuế": max(0.0, thu_nhap_tinh_thue),
        "Thuế TNCN": thue_tncn,
        "Tiền thực lãnh": thuc_lanh,
    }

# =====================================================
# 6. TẠO FILE EXCEL CÓ ĐỊNH DẠNG
# =====================================================
def tao_file_excel(df_chi_tiet):
    output = BytesIO()

    cot_tien = [
        "Lương cơ bản",
        "Tiền thưởng",
        "Phụ cấp tính BHXH",
        "Phụ cấp không tính BHXH",
        "Tổng thu nhập",
        "Bảo hiểm người lao động đóng",
        "Giảm trừ bản thân",
        "Giảm trừ người phụ thuộc",
        "Thu nhập tính thuế",
        "Thuế TNCN",
        "Tiền thực lãnh",
    ]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet 1: Bảng lương tổng hợp
        df_chi_tiet.to_excel(
            writer,
            sheet_name="Bang luong",
            index=False,
            startrow=2
        )

        ws = writer.sheets["Bang luong"]
        ws["A1"] = "BẢNG TỔNG HỢP LƯƠNG VÀ THUẾ TNCN"
        ws["A2"] = "Đơn vị tiền: triệu đồng/tháng"

        # Sheet 2: Chi tiết từng nhân viên
        df_chi_tiet.to_excel(
            writer,
            sheet_name="Chi tiet",
            index=False
        )

        # Định dạng cả hai sheet
        from openpyxl.styles import Font, PatternFill, Alignment

        for sheet_name in ["Bang luong", "Chi tiet"]:
            sheet = writer.sheets[sheet_name]
            header_row = 3 if sheet_name == "Bang luong" else 1

            for cell in sheet[header_row]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="1F4E78"
                )
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )

            sheet.freeze_panes = f"A{header_row + 1}"
            sheet.auto_filter.ref = (
                f"A{header_row}:{sheet.cell(header_row, sheet.max_column).coordinate}"
                f"{sheet.max_row}"
            )

            for col_cells in sheet.columns:
                col_letter = col_cells[0].column_letter
                max_length = max(
                    len(str(cell.value or ""))
                    for cell in col_cells
                )
                sheet.column_dimensions[col_letter].width = min(
                    max(max_length + 2, 14), 32
                )

            for row in sheet.iter_rows(
                min_row=header_row + 1,
                max_row=sheet.max_row
            ):
                for cell in row:
                    cell.alignment = Alignment(vertical="center")
                    if cell.column > 1 and cell.value is not None:
                        # Định dạng số tiền, không định dạng cột số người phụ thuộc
                        header = sheet.cell(header_row, cell.column).value
                        if header in cot_tien:
                            cell.number_format = '#,##0.000;[Red]-#,##0.000;0.000'

        # Thêm dòng tổng cộng vào cuối bảng lương
        ws = writer.sheets["Bang luong"]
        dong_tong = ws.max_row + 1
        ws.cell(dong_tong, 1, "TỔNG CỘNG")

        cot_tong_hop = [
            "Lương cơ bản",
            "Tiền thưởng",
            "Phụ cấp tính BHXH",
            "Phụ cấp không tính BHXH",
            "Tổng thu nhập",
            "Bảo hiểm người lao động đóng",
            "Thuế TNCN",
            "Tiền thực lãnh",
        ]

        # Tìm cột theo tiêu đề, tránh phụ thuộc thứ tự cột
        tieu_de = {
            ws.cell(3, col).value: col
            for col in range(1, ws.max_column + 1)
        }

        for ten_cot in cot_tong_hop:
            if ten_cot in tieu_de:
                col = tieu_de[ten_cot]
                cot_chu = ws.cell(3, col).column_letter
                ws.cell(
                    dong_tong,
                    col,
                    f"=SUM({cot_chu}4:{cot_chu}{dong_tong - 1})"
                )
                ws.cell(dong_tong, col).number_format = '#,##0.000'

        for cell in ws[dong_tong]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="D9EAF7"
            )

    output.seek(0)
    return output.getvalue()

# =====================================================
# 7. FONT TIẾNG VIỆT CHO PDF
# Tìm font Unicode có sẵn trên máy; nếu không có (Streamlit Cloud)
# thì tự tải DejaVu Sans về thư mục tạm. Nếu vẫn không được thì
# dùng font mặc định và bỏ dấu tiếng Việt.
# =====================================================
FONT_URL = "https://cdn.jsdelivr.net/npm/dejavu-fonts-ttf@2.37.3/ttf/"


@st.cache_resource(show_spinner=False)
def tim_font_unicode():
    ung_vien = [
        (THU_MUC_GOC / "fonts" / "DejaVuSans.ttf",
         THU_MUC_GOC / "fonts" / "DejaVuSans-Bold.ttf"),
        (Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
         Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")),
        (Path("C:/Windows/Fonts/arial.ttf"),
         Path("C:/Windows/Fonts/arialbd.ttf")),
    ]
    for thuong, dam in ung_vien:
        if thuong.exists() and dam.exists():
            return str(thuong), str(dam)

    thu_muc_tam = Path(tempfile.gettempdir()) / "font_phieu_luong"
    thu_muc_tam.mkdir(exist_ok=True)
    try:
        ket_qua = []
        for ten in ["DejaVuSans.ttf", "DejaVuSans-Bold.ttf"]:
            dich = thu_muc_tam / ten
            if not dich.exists():
                urllib.request.urlretrieve(FONT_URL + ten, dich)
            ket_qua.append(str(dich))
        return tuple(ket_qua)
    except Exception:
        return None

# =====================================================
# 8. TẠO PHIẾU LƯƠNG PDF CHO TỪNG NGƯỜI
# =====================================================
def tao_pdf_phieu_luong(nv, ky_luong, ten_cong_ty):
    font = tim_font_unicode()
    pdf = FPDF(format="A4")
    pdf.set_margins(18, 18, 18)
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()

    if font:
        pdf.add_font("VN", "", font[0])
        pdf.add_font("VN", "B", font[1])
        ten_font = "VN"

        def chu(s):
            return str(s)
    else:
        ten_font = "Helvetica"

        def chu(s):
            return bo_dau(s)

    rong = pdf.w - pdf.l_margin - pdf.r_margin
    xanh = (31, 78, 120)

    # Tiêu đề
    pdf.set_font(ten_font, "B", 11)
    pdf.set_text_color(*xanh)
    pdf.cell(rong, 6, chu(ten_cong_ty.upper()),
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)
    pdf.set_font(ten_font, "B", 18)
    pdf.cell(rong, 10, chu("PHIẾU LƯƠNG"), align="C",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font(ten_font, "", 11)
    pdf.set_text_color(90, 90, 90)
    pdf.cell(rong, 6, chu(f"Kỳ lương: {ky_luong}"), align="C",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(6)

    # Thông tin nhân viên
    pdf.set_text_color(0, 0, 0)
    for nhan, gia_tri in [
        ("Họ và tên", nv["Họ và tên"]),
        ("Email", nv["Email"]),
        ("Số người phụ thuộc", int(nv["Số người phụ thuộc"])),
    ]:
        pdf.set_font(ten_font, "", 11)
        pdf.cell(50, 7, chu(f"{nhan}:"))
        pdf.set_font(ten_font, "B", 11)
        pdf.multi_cell(rong - 50, 7, chu(gia_tri),
                       new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)

    # Bảng chi tiết
    cot_1, cot_2 = rong * 0.68, rong * 0.32
    pdf.set_fill_color(*xanh)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(ten_font, "B", 11)
    pdf.cell(cot_1, 9, chu("  Khoản mục"), border=1, fill=True)
    pdf.cell(cot_2, 9, chu("Số tiền (triệu đồng)  "), border=1, fill=True,
             align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    dong_bang = [
        ("Lương cơ bản", "Lương cơ bản", False),
        ("Tiền thưởng", "Tiền thưởng", False),
        ("Phụ cấp tính BHXH", "Phụ cấp tính BHXH", False),
        ("Phụ cấp không tính BHXH", "Phụ cấp không tính BHXH", False),
        ("TỔNG THU NHẬP", "Tổng thu nhập", True),
        ("Bảo hiểm người lao động đóng (10,5%)", "Bảo hiểm người lao động đóng", False),
        ("Giảm trừ bản thân", "Giảm trừ bản thân", False),
        ("Giảm trừ người phụ thuộc", "Giảm trừ người phụ thuộc", False),
        ("Thu nhập tính thuế", "Thu nhập tính thuế", False),
        ("Thuế TNCN", "Thuế TNCN", False),
    ]
    pdf.set_text_color(0, 0, 0)
    for i, (nhan, cot, dam) in enumerate(dong_bang):
        pdf.set_font(ten_font, "B" if dam else "", 11)
        to_nen = dam or i % 2 == 1
        pdf.set_fill_color(*((217, 234, 247) if dam else (245, 247, 250)))
        pdf.cell(cot_1, 8, chu(f"  {nhan}"), border=1, fill=to_nen)
        pdf.cell(cot_2, 8, dinh_dang_tien(float(nv[cot])) + "  ", border=1,
                 fill=to_nen, align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Thực lãnh
    pdf.set_fill_color(*xanh)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(ten_font, "B", 13)
    pdf.cell(cot_1, 11, chu("  TIỀN THỰC LÃNH"), border=1, fill=True)
    pdf.cell(cot_2, 11, dinh_dang_tien(float(nv["Tiền thực lãnh"])) + "  ",
             border=1, fill=True, align="R",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Ghi chú
    pdf.ln(6)
    pdf.set_text_color(110, 110, 110)
    pdf.set_font(ten_font, "", 9)
    pdf.multi_cell(rong, 5, chu(
        "Đơn vị tiền: triệu đồng/tháng. Tiền thực lãnh = Tổng thu nhập − "
        "Bảo hiểm − Thuế TNCN. Phiếu lương được tạo tự động ngày "
        f"{date.today():%d/%m/%Y}."
    ))

    return bytes(pdf.output())


def ten_file_pdf(stt, nv):
    return f"{stt:03d}_{ten_file_an_toan(nv['Họ và tên'])}.pdf"


def tao_zip_phieu_luong(df, ky_luong, ten_cong_ty, thanh_tien_do=None):
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (_, nv) in enumerate(df.iterrows(), start=1):
            zf.writestr(
                ten_file_pdf(i, nv),
                tao_pdf_phieu_luong(nv, ky_luong, ten_cong_ty)
            )
            if thanh_tien_do:
                thanh_tien_do.progress(i / len(df), text=f"Đang tạo {i}/{len(df)} phiếu lương...")
        # Danh sách email kèm tên file để tiện gửi phiếu lương
        danh_sach = pd.DataFrame({
            "STT": range(1, len(df) + 1),
            "Họ và tên": df["Họ và tên"].values,
            "Email": df["Email"].values,
            "File PDF": [ten_file_pdf(i, nv) for i, (_, nv) in enumerate(df.iterrows(), start=1)],
        })
        zf.writestr("danh_sach_email.csv", danh_sach.to_csv(index=False).encode("utf-8-sig"))
    output.seek(0)
    return output.getvalue()

# =====================================================
# 8b. GỬI PHIẾU LƯƠNG QUA GMAIL
# Dùng SMTP của Gmail với "Mật khẩu ứng dụng" (App Password),
# không dùng mật khẩu đăng nhập Gmail thông thường.
# =====================================================
GMAIL_SMTP_HOST = "smtp.gmail.com"
GMAIL_SMTP_PORT = 465


def lay_secret(ten):
    # Đọc từ Settings > Secrets trên Streamlit Cloud (hoặc .streamlit/secrets.toml)
    try:
        return str(st.secrets.get(ten, ""))
    except Exception:
        return ""


def dien_mau(mau, nv, ky_luong, ten_cong_ty):
    # Thay các biến {ho_ten}, {email}, {ky_luong}, {cong_ty}, {thuc_lanh}
    gia_tri = {
        "ho_ten": nv["Họ và tên"],
        "email": nv["Email"],
        "ky_luong": ky_luong,
        "cong_ty": ten_cong_ty,
        "thuc_lanh": dinh_dang_tien(float(nv["Tiền thực lãnh"])),
    }
    for khoa, gt in gia_tri.items():
        mau = mau.replace("{" + khoa + "}", str(gt))
    return mau


def tao_email_phieu_luong(nv, nguoi_gui, ten_nguoi_gui, tieu_de, noi_dung,
                          ky_luong, ten_cong_ty, stt):
    thu = EmailMessage()
    thu["From"] = f"{ten_nguoi_gui} <{nguoi_gui}>" if ten_nguoi_gui else nguoi_gui
    thu["To"] = str(nv["Email"]).strip()
    thu["Subject"] = dien_mau(tieu_de, nv, ky_luong, ten_cong_ty)
    thu.set_content(dien_mau(noi_dung, nv, ky_luong, ten_cong_ty))
    thu.add_attachment(
        tao_pdf_phieu_luong(nv, ky_luong, ten_cong_ty),
        maintype="application",
        subtype="pdf",
        filename=ten_file_pdf(stt, nv),
    )
    return thu


def thong_bao_loi_email(loi):
    # Đổi lỗi kỹ thuật thành thông báo dễ hiểu cho người dùng
    if isinstance(loi, smtplib.SMTPRecipientsRefused):
        return "Email không tồn tại hoặc bị máy chủ nhận từ chối"
    if isinstance(loi, smtplib.SMTPAuthenticationError):
        return "Sai Gmail người gửi hoặc Mật khẩu ứng dụng"
    if isinstance(loi, smtplib.SMTPSenderRefused):
        return "Gmail từ chối người gửi (có thể đã vượt giới hạn gửi trong ngày)"
    if isinstance(loi, smtplib.SMTPServerDisconnected):
        return "Gmail đã ngắt kết nối (có thể đã vượt giới hạn gửi trong ngày)"
    if isinstance(loi, smtplib.SMTPDataError):
        return "Gmail từ chối nội dung email"
    if isinstance(loi, (TimeoutError, OSError)):
        return "Không kết nối được tới Gmail, kiểm tra lại kết nối mạng"
    return "Không gửi được email"


# =====================================================
# 8c. KIỂM TRA EMAIL TRƯỚC KHI GỬI
# Kiểm tra định dạng, tên miền có tồn tại và có nhận email (bản ghi MX)
# và tên miền gõ nhầm (gmial.com, yahooo.com...).
# Không thể xác minh chắc chắn từng hộp thư: Gmail, Outlook... không
# cho phép dò hộp thư từ bên ngoài.
# =====================================================
TEN_MIEN_PHO_BIEN = [
    "gmail.com", "yahoo.com", "yahoo.com.vn", "outlook.com", "hotmail.com",
    "icloud.com", "live.com", "msn.com", "fpt.com.vn", "vnpt.vn", "viettel.com.vn",
]


@st.cache_data(ttl=3600, show_spinner=False)
def kiem_tra_ten_mien(ten_mien):
    # Trả về: "ok", "khong_ton_tai", "khong_nhan_mail", "khong_ro"
    try:
        ban_ghi = dns.resolver.resolve(ten_mien, "MX", lifetime=5)
        # Null MX (RFC 7505): tên miền khai báo không nhận email
        if all(str(r.exchange) == "." for r in ban_ghi):
            return "khong_nhan_mail"
        return "ok"
    except dns.resolver.NXDOMAIN:
        return "khong_ton_tai"
    except dns.resolver.NoAnswer:
        # Không có MX thì máy chủ gửi sẽ thử bản ghi A của tên miền
        try:
            dns.resolver.resolve(ten_mien, "A", lifetime=5)
            return "ok"
        except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
            return "khong_nhan_mail"
        except Exception:
            return "khong_ro"
    except Exception:
        return "khong_ro"


def kiem_tra_email(email):
    # Trả về (có gửi được không, ghi chú)
    email = str(email).strip()
    if not email_hop_le(email):
        return False, "Sai định dạng email"

    ten_mien = email.rsplit("@", 1)[1].lower()
    if ten_mien not in TEN_MIEN_PHO_BIEN:
        goi_y = difflib.get_close_matches(ten_mien, TEN_MIEN_PHO_BIEN, n=1, cutoff=0.85)
        if goi_y:
            return False, f"Nghi gõ nhầm tên miền, có phải là @{goi_y[0]}?"

    trang_thai = kiem_tra_ten_mien(ten_mien)
    if trang_thai == "khong_ton_tai":
        return False, f"Tên miền @{ten_mien} không tồn tại"
    if trang_thai == "khong_nhan_mail":
        return False, f"Tên miền @{ten_mien} không nhận email"
    if trang_thai == "khong_ro":
        return True, "Không kiểm tra được tên miền (vẫn gửi)"
    return True, ""


def kiem_tra_danh_sach_email(danh_sach, tat_ca_email, thanh_tien_do=None):
    # danh_sach: list (stt, nv) cần gửi; tat_ca_email: email của toàn bộ nhân viên
    # (đếm trùng trên toàn bộ danh sách, kể cả người không được chọn gửi)
    dem_email = pd.Series([str(e).strip().lower() for e in tat_ca_email]).value_counts()
    ket_qua = []
    for i, (stt, nv) in enumerate(danh_sach, start=1):
        gui_duoc, ghi_chu = kiem_tra_email(nv["Email"])
        # Email trùng: chặn để tránh một người nhận phiếu lương của người khác
        if gui_duoc and dem_email[str(nv["Email"]).strip().lower()] > 1:
            gui_duoc, ghi_chu = False, "Email bị trùng với nhân viên khác"
        ket_qua.append({
            "STT": stt,
            "Họ và tên": nv["Họ và tên"],
            "Email": nv["Email"],
            "Kết quả kiểm tra": "✅ Hợp lệ" if gui_duoc else "❌ Không hợp lệ",
            "Ghi chú": ghi_chu,
        })
        if thanh_tien_do:
            thanh_tien_do.progress(i / len(danh_sach),
                                   text=f"Đang kiểm tra {i}/{len(danh_sach)} email...")
    return pd.DataFrame(ket_qua)


def gui_phieu_luong_gmail(danh_sach, nguoi_gui, mat_khau_ung_dung, ten_nguoi_gui,
                          tieu_de, noi_dung, ky_luong, ten_cong_ty,
                          thanh_tien_do=None, nghi_giay=1.0):
    # danh_sach: list (stt, nv). Trả về DataFrame kết quả gửi từng người
    ket_qua = []
    ngu_canh = ssl.create_default_context()
    with smtplib.SMTP_SSL(GMAIL_SMTP_HOST, GMAIL_SMTP_PORT, context=ngu_canh, timeout=60) as smtp:
        smtp.login(nguoi_gui, mat_khau_ung_dung.replace(" ", ""))
        for i, (stt, nv) in enumerate(danh_sach, start=1):
            try:
                smtp.send_message(tao_email_phieu_luong(
                    nv, nguoi_gui, ten_nguoi_gui, tieu_de, noi_dung,
                    ky_luong, ten_cong_ty, stt
                ))
                trang_thai, loi = "✅ Đã gửi", ""
            except smtplib.SMTPServerDisconnected as e:
                # Gmail ngắt kết nối (thường do vượt giới hạn) -> dừng
                ket_qua.append({"STT": stt, "Họ và tên": nv["Họ và tên"], "Email": nv["Email"],
                                "Trạng thái": "❌ Không gửi được", "Chi tiết": thong_bao_loi_email(e)})
                break
            except Exception as e:
                trang_thai, loi = "❌ Không gửi được", thong_bao_loi_email(e)
            ket_qua.append({"STT": stt, "Họ và tên": nv["Họ và tên"], "Email": nv["Email"],
                            "Trạng thái": trang_thai, "Chi tiết": loi})
            if thanh_tien_do:
                thanh_tien_do.progress(i / len(danh_sach),
                                       text=f"Đã gửi {i}/{len(danh_sach)} email...")
            if i < len(danh_sach):
                time.sleep(nghi_giay)
    return pd.DataFrame(ket_qua)

# =====================================================
# 8d. KIỂM TRA DỮ LIỆU NHẬP
# =====================================================
COT_SO = [
    "Lương cơ bản",
    "Tiền thưởng",
    "Phụ cấp tính BHXH",
    "Phụ cấp không tính BHXH",
    "Số người phụ thuộc",
]

GIA_TRI_RONG = {"", "none", "nan", "null", "<na>"}


def la_o_trong(gia_tri):
    return gia_tri is None or pd.isna(gia_tri) or str(gia_tri).strip().lower() in GIA_TRI_RONG


def chuyen_so(gia_tri):
    # Trả về số thực, hoặc None nếu không phải số (chấp nhận "1,5" kiểu Việt Nam).
    # Ô để trống được tính là 0
    if la_o_trong(gia_tri):
        return 0.0
    if isinstance(gia_tri, (int, float)) and not isinstance(gia_tri, bool):
        return float(gia_tri)
    chuoi = str(gia_tri).strip().replace(" ", "")
    if "," in chuoi and "." not in chuoi:
        chuoi = chuoi.replace(",", ".")
    try:
        return float(chuoi)
    except ValueError:
        return None


def kiem_tra_du_lieu_nhap(df):
    # Trả về danh sách lỗi theo từng ô: dòng, họ tên, cột, nội dung lỗi
    loi = []
    for dong, (_, nv) in enumerate(df.iterrows(), start=1):
        ho_ten = "" if la_o_trong(nv["Họ và tên"]) else str(nv["Họ và tên"]).strip()

        def them_loi(cot, noi_dung):
            loi.append({"Dòng": dong, "Họ và tên": ho_ten or "(chưa có tên)",
                        "Cột": cot, "Lỗi": noi_dung})

        if not ho_ten:
            them_loi("Họ và tên", "Chưa nhập dữ liệu, phải nhập họ tên")

        if la_o_trong(nv["Email"]):
            them_loi("Email", "Chưa nhập dữ liệu, phải nhập email")
        elif not email_hop_le(nv["Email"]):
            them_loi("Email", f"Email sai định dạng: '{str(nv['Email']).strip()}'")

        for cot in COT_SO:
            gia_tri = nv[cot]
            so = chuyen_so(gia_tri)
            if so is None:
                them_loi(cot, f"Phải nhập số, không nhập chữ: '{str(gia_tri).strip()}'")
            elif so < 0:
                them_loi(cot, "Không được nhập số âm")
            elif cot == "Số người phụ thuộc" and so != int(so):
                them_loi(cot, "Phải nhập số nguyên (0, 1, 2...)")
    return loi

# =====================================================
# 9. GIAO DIỆN ỨNG DỤNG
# =====================================================
st.title("💰 PHẦN MỀM TÍNH TIỀN LƯƠNG CHO TOÀN BỘ NHÂN VIÊN")
st.subheader("HVTH: NGUYỄN THỊ HUYỀN TRÂN")

st.caption(
    "Tính đồng thời tiền lương, bảo hiểm, thuế TNCN và tiền thực lãnh "
    "cho tối đa 500 người lao động, xuất phiếu lương PDF cho từng người."
)

st.info(
    "Đơn vị nhập liệu là triệu đồng/tháng. "
    "Ứng dụng sử dụng các mức giảm trừ và biểu thuế từ code ban đầu. "
    "Vui lòng đối chiếu quy định pháp luật hiện hành trước khi dùng để trả lương thực tế."
)

# =====================================================
# 10. CHỌN SỐ LƯỢNG NHÂN VIÊN
# =====================================================
st.subheader("1. Chọn số lượng người lao động")

so_nhan_vien = st.number_input(
    "Số người lao động cần tính",
    min_value=1,
    max_value=500,
    value=10,
    step=1
)

st.caption(
    f"Bạn đang chọn tính lương cho {so_nhan_vien} người lao động."
)

# =====================================================
# 11. NHẬP DỮ LIỆU NHÂN VIÊN
# =====================================================
st.subheader("2. Nhập thông tin lương")

st.markdown(
    "Nhập trực tiếp trên bảng bên dưới. "
    "Có thể sao chép dữ liệu từ Excel và dán vào bảng. "
    "Ô số để trống được tính là 0."
)

COT_NHAP = [
    "Họ và tên",
    "Email",
    "Lương cơ bản",
    "Tiền thưởng",
    "Phụ cấp tính BHXH",
    "Phụ cấp không tính BHXH",
    "Số người phụ thuộc",
]

if "du_lieu_nhan_vien" not in st.session_state:
    st.session_state.du_lieu_nhan_vien = pd.DataFrame(
        {
            "Họ và tên": [f"Nhân viên {i + 1}" for i in range(500)],
            "Email": [f"nhanvien{i + 1}@congty.vn" for i in range(500)],
            "Lương cơ bản": [15.0] * 500,
            "Tiền thưởng": [2.0] * 500,
            "Phụ cấp tính BHXH": [1.0] * 500,
            "Phụ cấp không tính BHXH": [0.5] * 500,
            "Số người phụ thuộc": [1] * 500,
        }
    )
elif "Email" not in st.session_state.du_lieu_nhan_vien.columns:
    # Phiên cũ chưa có cột Email
    st.session_state.du_lieu_nhan_vien.insert(1, "Email", "")

# Các cột số luôn để kiểu số thực, để ô bị xóa (trống) không gây lỗi kiểu dữ liệu
st.session_state.du_lieu_nhan_vien[COT_SO] = (
    st.session_state.du_lieu_nhan_vien[COT_SO].astype(float)
)

# Đổi phiên bản bảng nhập để vẽ lại bảng khi tự điền 0 vào ô trống
if "phien_bang_nhap" not in st.session_state:
    st.session_state.phien_bang_nhap = 0

# Hiển thị số dòng theo lựa chọn
du_lieu_goc = st.session_state.du_lieu_nhan_vien.iloc[:int(so_nhan_vien)][COT_NHAP].copy()

du_lieu_nhap = st.data_editor(
    du_lieu_goc,
    width="stretch",
    hide_index=True,
    num_rows="fixed",
    key=f"bang_nhap_{int(so_nhan_vien)}_{st.session_state.phien_bang_nhap}",
    column_config={
        "Họ và tên": st.column_config.TextColumn(
            "Họ và tên",
            required=True,
            width="medium"
        ),
        "Email": st.column_config.TextColumn(
            "Email",
            required=True,
            width="medium",
            validate=MAU_EMAIL.pattern
        ),
        "Lương cơ bản": st.column_config.NumberColumn(
            "Lương cơ bản",
            min_value=0.0,
            step=0.1,
            format="%.3f"
        ),
        "Tiền thưởng": st.column_config.NumberColumn(
            "Tiền thưởng",
            min_value=0.0,
            step=0.1,
            format="%.3f"
        ),
        "Phụ cấp tính BHXH": st.column_config.NumberColumn(
            "Phụ cấp tính BHXH",
            min_value=0.0,
            step=0.1,
            format="%.3f"
        ),
        "Phụ cấp không tính BHXH": st.column_config.NumberColumn(
            "Phụ cấp không tính BHXH",
            min_value=0.0,
            step=0.1,
            format="%.3f"
        ),
        "Số người phụ thuộc": st.column_config.NumberColumn(
            "Số người phụ thuộc",
            min_value=0,
            max_value=100,
            step=1,
            format="%d"
        ),
    }
)

# Ô số bị xóa (trống) -> tự điền 0
du_lieu_nhap = du_lieu_nhap.reset_index(drop=True)
co_o_trong = du_lieu_nhap[COT_SO].isna().any().any()
du_lieu_nhap[COT_SO] = du_lieu_nhap[COT_SO].astype(float).fillna(0.0)

# Lưu dữ liệu vừa nhập để giữ lại khi đổi số lượng
du_lieu_cap_nhat = st.session_state.du_lieu_nhan_vien.copy()
for cot in COT_NHAP:
    du_lieu_cap_nhat.loc[:int(so_nhan_vien) - 1, cot] = du_lieu_nhap[cot].to_numpy()
st.session_state.du_lieu_nhan_vien = du_lieu_cap_nhat

if co_o_trong:
    # Vẽ lại bảng để ô vừa xóa hiển thị số 0
    st.session_state.phien_bang_nhap += 1
    st.rerun()

# =====================================================
# 12. TÍNH TOÁN TOÀN BỘ NHÂN VIÊN
# =====================================================
st.divider()

if st.button(
    "🧮 TÍNH LƯƠNG TOÀN BỘ NHÂN VIÊN",
    type="primary",
    width="stretch"
):
    loi_nhap_lieu = kiem_tra_du_lieu_nhap(du_lieu_nhap)

    if loi_nhap_lieu:
        st.error(
            f"❌ Có {len(loi_nhap_lieu)} ô nhập liệu chưa đúng. "
            "Vui lòng nhập / sửa dữ liệu ở bảng trên rồi bấm tính lương lại:"
        )
        st.dataframe(
            pd.DataFrame(loi_nhap_lieu),
            width="stretch",
            hide_index=True
        )
        st.stop()

    # Dữ liệu đã kiểm tra: chuyển các cột số về kiểu số để tính toán
    du_lieu_nhap = du_lieu_nhap.copy()
    for cot in COT_SO:
        du_lieu_nhap[cot] = du_lieu_nhap[cot].map(chuyen_so)

    ket_qua = []

    for _, nv in du_lieu_nhap.iterrows():
        ket_qua_nv = tinh_luong(nv)

        dong = {
            "Họ và tên": str(nv["Họ và tên"]).strip(),
            "Email": str(nv["Email"]).strip(),
            "Lương cơ bản": float(nv["Lương cơ bản"]),
            "Tiền thưởng": float(nv["Tiền thưởng"]),
            "Phụ cấp tính BHXH": float(nv["Phụ cấp tính BHXH"]),
            "Phụ cấp không tính BHXH": float(nv["Phụ cấp không tính BHXH"]),
            "Số người phụ thuộc": int(nv["Số người phụ thuộc"]),
        }

        dong.update(ket_qua_nv)
        ket_qua.append(dong)

    df_ket_qua = pd.DataFrame(ket_qua)

    # Lưu kết quả để có thể tải file sau khi Streamlit chạy lại
    st.session_state.ket_qua_luong = df_ket_qua
    # Kết quả mới -> file ZIP cũ không còn đúng
    st.session_state.pop("file_zip_pdf", None)
    st.session_state.pop("ket_qua_gui_email", None)
    st.session_state.pop("ket_qua_kiem_tra_email", None)

# =====================================================
# 13. HIỂN THỊ KẾT QUẢ
# =====================================================
if "ket_qua_luong" in st.session_state:
    df_ket_qua = st.session_state.ket_qua_luong

    st.subheader("3. Kết quả tính lương")

    tong_thu_nhap_all = df_ket_qua["Tổng thu nhập"].sum()
    tong_bao_hiem_all = df_ket_qua["Bảo hiểm người lao động đóng"].sum()
    tong_thue_all = df_ket_qua["Thuế TNCN"].sum()
    tong_thuc_lanh_all = df_ket_qua["Tiền thực lãnh"].sum()

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Số nhân viên",
        f"{len(df_ket_qua):,}"
    )

    c2.metric(
        "Tổng thu nhập",
        f"{dinh_dang_tien(tong_thu_nhap_all)} triệu"
    )

    c3.metric(
        "Tổng thuế TNCN",
        f"{dinh_dang_tien(tong_thue_all)} triệu"
    )

    c4.metric(
        "Tổng tiền thực lãnh",
        f"{dinh_dang_tien(tong_thuc_lanh_all)} triệu"
    )

    st.markdown("#### Bảng lương chi tiết")

    cot_hien_thi = [
        "Họ và tên",
        "Email",
        "Lương cơ bản",
        "Tiền thưởng",
        "Phụ cấp tính BHXH",
        "Phụ cấp không tính BHXH",
        "Số người phụ thuộc",
        "Tổng thu nhập",
        "Bảo hiểm người lao động đóng",
        "Thu nhập tính thuế",
        "Thuế TNCN",
        "Tiền thực lãnh",
    ]

    df_hien_thi = df_ket_qua[cot_hien_thi].copy()

    cot_tien_hien_thi = [
        "Lương cơ bản",
        "Tiền thưởng",
        "Phụ cấp tính BHXH",
        "Phụ cấp không tính BHXH",
        "Tổng thu nhập",
        "Bảo hiểm người lao động đóng",
        "Thu nhập tính thuế",
        "Thuế TNCN",
        "Tiền thực lãnh",
    ]

    st.dataframe(
        df_hien_thi,
        width="stretch",
        hide_index=True,
        column_config={
            cot: st.column_config.NumberColumn(
                cot,
                format="%.3f"
            )
            for cot in cot_tien_hien_thi
        }
    )

    # Dòng tổng cộng trên giao diện
    st.markdown("#### Tổng cộng toàn bộ nhân viên")

    df_tong = pd.DataFrame([{
        "Họ và tên": "TỔNG CỘNG",
        "Lương cơ bản": df_ket_qua["Lương cơ bản"].sum(),
        "Tiền thưởng": df_ket_qua["Tiền thưởng"].sum(),
        "Phụ cấp tính BHXH": df_ket_qua["Phụ cấp tính BHXH"].sum(),
        "Phụ cấp không tính BHXH": df_ket_qua["Phụ cấp không tính BHXH"].sum(),
        "Số người phụ thuộc": df_ket_qua["Số người phụ thuộc"].sum(),
        "Tổng thu nhập": tong_thu_nhap_all,
        "Bảo hiểm người lao động đóng": tong_bao_hiem_all,
        "Thu nhập tính thuế": df_ket_qua["Thu nhập tính thuế"].sum(),
        "Thuế TNCN": tong_thue_all,
        "Tiền thực lãnh": tong_thuc_lanh_all,
    }])

    st.dataframe(
        df_tong,
        width="stretch",
        hide_index=True,
        column_config={
            cot: st.column_config.NumberColumn(
                cot,
                format="%.3f"
            )
            for cot in cot_tien_hien_thi
        }
    )

    # Xuất Excel
    st.subheader("4. Xuất bảng lương Excel")

    file_excel = tao_file_excel(df_ket_qua)

    st.download_button(
        label="📥 TẢI FILE EXCEL BẢNG LƯƠNG (.xlsx)",
        data=file_excel,
        file_name="bang_luong_va_thue_TNCN.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch"
    )

    # Xuất thêm CSV nếu cần
    csv_data = df_ket_qua.to_csv(
        index=False
    ).encode("utf-8-sig")

    st.download_button(
        label="📥 TẢI FILE CSV",
        data=csv_data,
        file_name="bang_luong_va_thue_TNCN.csv",
        mime="text/csv",
        width="stretch"
    )

    # =================================================
    # 14. XUẤT PHIẾU LƯƠNG PDF CHO TỪNG NGƯỜI
    # =================================================
    st.subheader("5. Xuất phiếu lương PDF cho từng người")

    c1, c2 = st.columns(2)
    ten_cong_ty = c1.text_input("Tên công ty / đơn vị", value="CÔNG TY ABC")
    ky_luong = c2.text_input("Kỳ lương", value=f"Tháng {date.today():%m/%Y}")

    if tim_font_unicode() is None:
        st.warning(
            "Không tìm thấy font tiếng Việt và không tải được font từ Internet, "
            "PDF sẽ hiển thị tiếng Việt không dấu. Có thể đặt DejaVuSans.ttf và "
            "DejaVuSans-Bold.ttf vào thư mục 'fonts' cạnh file app."
        )

    tab_mot, tab_tat_ca, tab_email = st.tabs(
        ["📄 Từng người", "🗂️ Tất cả (ZIP)", "✉️ Gửi email (Gmail)"]
    )

    with tab_mot:
        chi_so = st.selectbox(
            "Chọn nhân viên",
            options=list(range(len(df_ket_qua))),
            format_func=lambda i: (
                f"{i + 1}. {df_ket_qua.iloc[i]['Họ và tên']} "
                f"— {df_ket_qua.iloc[i]['Email']}"
            ),
        )
        nv_chon = df_ket_qua.iloc[chi_so]

        st.download_button(
            label=f"📥 TẢI PHIẾU LƯƠNG PDF — {nv_chon['Họ và tên']}",
            data=tao_pdf_phieu_luong(nv_chon, ky_luong, ten_cong_ty),
            file_name=ten_file_pdf(chi_so + 1, nv_chon),
            mime="application/pdf",
            width="stretch"
        )

    with tab_tat_ca:
        st.caption(
            "Mỗi nhân viên một file PDF, đóng gói trong 1 file ZIP "
            "kèm danh_sach_email.csv (họ tên, email, tên file) để gửi phiếu lương."
        )

        if st.button("⚙️ TẠO PHIẾU LƯƠNG PDF CHO TẤT CẢ", width="stretch"):
            tien_do = st.progress(0.0, text="Đang tạo phiếu lương...")
            st.session_state.file_zip_pdf = tao_zip_phieu_luong(
                df_ket_qua, ky_luong, ten_cong_ty, tien_do
            )
            tien_do.empty()

        if "file_zip_pdf" in st.session_state:
            st.download_button(
                label=f"📥 TẢI FILE ZIP ({len(df_ket_qua)} PHIẾU LƯƠNG PDF)",
                data=st.session_state.file_zip_pdf,
                file_name="phieu_luong_pdf.zip",
                mime="application/zip",
                width="stretch"
            )

    with tab_email:
        st.caption(
            "Gửi phiếu lương PDF đính kèm tới email của từng nhân viên qua Gmail. "
            "Cần bật Xác minh 2 bước và tạo **Mật khẩu ứng dụng** (16 ký tự) tại "
            "https://myaccount.google.com/apppasswords — không dùng mật khẩu Gmail thường."
        )

        gmail_secret = lay_secret("GMAIL_USER")
        mat_khau_secret = lay_secret("GMAIL_APP_PASSWORD")

        if gmail_secret and mat_khau_secret:
            st.success(f"Đang dùng tài khoản Gmail từ Secrets: {gmail_secret}")
            nguoi_gui, mat_khau = gmail_secret, mat_khau_secret
        else:
            c1, c2 = st.columns(2)
            nguoi_gui = c1.text_input("Gmail người gửi", placeholder="tencongty@gmail.com")
            mat_khau = c2.text_input("Mật khẩu ứng dụng", type="password",
                                     placeholder="xxxx xxxx xxxx xxxx")

        ten_nguoi_gui = st.text_input("Tên hiển thị người gửi", value=f"Phòng Nhân sự - {ten_cong_ty}")
        tieu_de = st.text_input("Tiêu đề email", value="[{cong_ty}] Phiếu lương {ky_luong} - {ho_ten}")
        noi_dung = st.text_area(
            "Nội dung email",
            height=180,
            value=(
                "Kính gửi Anh/Chị {ho_ten},\n\n"
                "Phòng Nhân sự {cong_ty} gửi Anh/Chị thông tin chi tiết về thu nhập và các khoản trích theo lương kỳ {ky_luong} (file PDF đính kèm).\n"
                "Tiền thực lãnh: {thuc_lanh} triệu đồng.\n\n"
                "Mọi ý kiến phản hồi hoặc thắc mắc liên quan đến kỳ lương này, Anh/Chị vui lòng liên hệ Bộ phận Nhân sự để được giải quyết ổn thỏa.\n\n"
                "Trân trọng cảm ơn,\nPhòng Nhân sự"
            ),
        )
        st.caption(
            "Biến có thể dùng: {ho_ten}, {email}, {ky_luong}, {cong_ty}, {thuc_lanh}"
        )

        pham_vi = st.radio(
            "Gửi cho",
            ["Tất cả nhân viên", "Chọn nhân viên"],
            horizontal=True,
        )
        if pham_vi == "Chọn nhân viên":
            chi_so_gui = st.multiselect(
                "Nhân viên nhận email",
                options=list(range(len(df_ket_qua))),
                format_func=lambda i: (
                    f"{i + 1}. {df_ket_qua.iloc[i]['Họ và tên']} "
                    f"— {df_ket_qua.iloc[i]['Email']}"
                ),
            )
        else:
            chi_so_gui = list(range(len(df_ket_qua)))

        st.caption(
            f"Sẽ gửi {len(chi_so_gui)} email. Gmail cá nhân giới hạn khoảng "
            "500 email/ngày (Google Workspace khoảng 2.000)."
        )

        if st.button("🔍 KIỂM TRA EMAIL TRƯỚC KHI GỬI", width="stretch",
                     disabled=not chi_so_gui):
            tien_do = st.progress(0.0, text="Đang kiểm tra email...")
            st.session_state.ket_qua_kiem_tra_email = kiem_tra_danh_sach_email(
                [(i + 1, df_ket_qua.iloc[i]) for i in chi_so_gui], df_ket_qua["Email"], tien_do
            )
            tien_do.empty()

        if "ket_qua_kiem_tra_email" in st.session_state:
            df_kt = st.session_state.ket_qua_kiem_tra_email
            df_kt_loi = df_kt[df_kt["Kết quả kiểm tra"] != "✅ Hợp lệ"]
            if df_kt_loi.empty:
                st.success(f"✅ Tất cả {len(df_kt)} email đều hợp lệ.")
            else:
                st.error(
                    f"❌ {len(df_kt_loi)}/{len(df_kt)} email không hợp lệ và sẽ không được gửi. "
                    "Hãy sửa email ở bảng nhập liệu (mục 2), bấm tính lương lại rồi kiểm tra lại."
                )
                st.dataframe(df_kt_loi, width="stretch", hide_index=True)
            df_kt_luu_y = df_kt[(df_kt["Kết quả kiểm tra"] == "✅ Hợp lệ") & (df_kt["Ghi chú"] != "")]
            if not df_kt_luu_y.empty:
                st.warning(f"{len(df_kt_luu_y)} email cần lưu ý (vẫn được gửi):")
                st.dataframe(df_kt_luu_y, width="stretch", hide_index=True)
            st.caption(
                "Kiểm tra được định dạng, tên miền có tồn tại / có nhận email và tên miền gõ nhầm. "
                "Không thể xác minh chắc chắn từng hộp thư (ví dụ sai tên trước dấu @ của Gmail) "
                "vì các nhà cung cấp email không cho phép dò hộp thư từ bên ngoài."
            )
        xac_nhan = st.checkbox(
            f"Tôi đã kiểm tra số liệu và xác nhận gửi {len(chi_so_gui)} email phiếu lương"
        )

        if st.button("📤 GỬI EMAIL PHIẾU LƯƠNG", type="primary", width="stretch",
                     disabled=not (xac_nhan and chi_so_gui)):
            if not email_hop_le(nguoi_gui) or not mat_khau.strip():
                st.error("Vui lòng nhập Gmail người gửi và Mật khẩu ứng dụng.")
            else:
                # Luôn kiểm tra lại email ngay trước khi gửi, bỏ qua email không hợp lệ
                tien_do = st.progress(0.0, text="Đang kiểm tra email...")
                danh_sach_gui = [(i + 1, df_ket_qua.iloc[i]) for i in chi_so_gui]
                df_kt = kiem_tra_danh_sach_email(danh_sach_gui, df_ket_qua["Email"], tien_do)
                st.session_state.ket_qua_kiem_tra_email = df_kt
                hop_le = (df_kt["Kết quả kiểm tra"] == "✅ Hợp lệ").tolist()
                df_bo_qua = df_kt[[not h for h in hop_le]].rename(
                    columns={"Kết quả kiểm tra": "Trạng thái", "Ghi chú": "Chi tiết"}
                ).assign(**{"Trạng thái": "❌ Không gửi được"})
                danh_sach_gui = [d for d, h in zip(danh_sach_gui, hop_le) if h]

                try:
                    df_da_gui = pd.DataFrame()
                    if danh_sach_gui:
                        tien_do.progress(0.0, text="Đang kết nối Gmail...")
                        df_da_gui = gui_phieu_luong_gmail(
                            danh_sach_gui,
                            nguoi_gui.strip(), mat_khau, ten_nguoi_gui,
                            tieu_de, noi_dung, ky_luong, ten_cong_ty, tien_do,
                        )
                    st.session_state.ket_qua_gui_email = (
                        pd.concat([df_da_gui, df_bo_qua], ignore_index=True)
                        .sort_values("STT", ignore_index=True)
                    )
                except smtplib.SMTPAuthenticationError:
                    st.session_state.pop("ket_qua_gui_email", None)
                    st.error(
                        "❌ Không gửi được email: đăng nhập Gmail thất bại. Kiểm tra lại "
                        "địa chỉ Gmail và Mật khẩu ứng dụng (không phải mật khẩu Gmail thường)."
                    )
                except Exception as e:
                    st.session_state.pop("ket_qua_gui_email", None)
                    st.error(f"❌ Không gửi được email: {thong_bao_loi_email(e)}.")
                tien_do.empty()

        if "ket_qua_gui_email" in st.session_state:
            df_gui = st.session_state.ket_qua_gui_email
            so_thanh_cong = (df_gui["Trạng thái"] == "✅ Đã gửi").sum()
            df_loi = df_gui[df_gui["Trạng thái"] != "✅ Đã gửi"]

            if so_thanh_cong == 0:
                st.error(f"❌ Không gửi được email nào ({len(df_gui)} email).")
            elif not df_loi.empty:
                st.warning(
                    f"Gửi thành công {so_thanh_cong}/{len(df_gui)} email. "
                    f"{len(df_loi)} email không gửi được:"
                )
            else:
                st.success(f"✅ Đã gửi thành công {so_thanh_cong}/{len(df_gui)} email.")

            for _, dong in df_loi.head(20).iterrows():
                st.markdown(f"- **{dong['Họ và tên']}** ({dong['Email']}): {dong['Chi tiết']}")
            if len(df_loi) > 20:
                st.markdown(f"- ... và {len(df_loi) - 20} email khác (xem nhật ký bên dưới)")

            st.caption(
                "Lưu ý: Gmail thường vẫn nhận gửi tới địa chỉ không tồn tại, sau đó mới "
                "trả về thư báo lỗi (Mail Delivery Subsystem) vào hộp thư của Gmail người gửi. "
                "Hãy kiểm tra hộp thư đó sau khi gửi."
            )
            st.dataframe(df_gui, width="stretch", hide_index=True)
            st.download_button(
                label="📥 TẢI NHẬT KÝ GỬI EMAIL (.csv)",
                data=df_gui.to_csv(index=False).encode("utf-8-sig"),
                file_name="nhat_ky_gui_email.csv",
                mime="text/csv",
                width="stretch",
            )

# =====================================================
# 15. GIẢI THÍCH VÀ LƯU Ý
# =====================================================
with st.expander("📘 Công thức tính lương"):
    st.markdown("""
    **1. Tổng thu nhập**
    Lương cơ bản + Tiền thưởng + Phụ cấp tính BHXH
    + Phụ cấp không tính BHXH.

    **2. Bảo hiểm người lao động đóng**
    (Lương cơ bản + Phụ cấp tính BHXH) × 10,5%.

    **3. Thu nhập tính thuế**
    Tổng thu nhập − 15,5 − (6,2 × số người phụ thuộc) − Bảo hiểm.

    **4. Thuế TNCN**
    Áp dụng công thức thuế lũy tiến được giữ lại từ code ban đầu.
    Nếu thu nhập tính thuế không dương, thuế được tính bằng 0.

    **5. Tiền thực lãnh**
    Tổng thu nhập − Bảo hiểm − Thuế TNCN.

    Đơn vị của các khoản tiền là triệu đồng/tháng.
    """)

st.divider()

st.caption(
    "Phần mềm tính theo các giả định được cài đặt trong mã nguồn. "
    "Cần đối chiếu quy định thuế, bảo hiểm và giảm trừ hiện hành trước khi "
    "dùng làm bảng lương chính thức."
)
