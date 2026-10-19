import os
import re
import ssl
import smtplib
import zipfile
from io import BytesIO
from email.message import EmailMessage

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =========================================================
# 1. CẤU HÌNH ỨNG DỤNG
# =========================================================

st.set_page_config(
    page_title="Quản lý lương và thuế TNCN",
    page_icon="💰",
    layout="wide"
)

st.title("💰 QUẢN LÝ LƯƠNG VÀ THUẾ TNCN - NGUYỄN CHÍ THANH")
st.caption(
    "Tính lương tối đa 500 nhân viên, xuất Excel và gửi "
    "phiếu lương riêng qua email."
)

st.warning(
    "Ứng dụng giữ công thức thuế và mức giảm trừ theo code ban đầu. "
    "Cần cập nhật quy định pháp luật hiện hành trước khi dùng để "
    "tính lương chính thức."
)


# =========================================================
# 2. HÀM TÍNH THUẾ
# =========================================================

def tinh_thue_tncn(thu_nhap_tinh_thue):
    # Giữ biểu thuế 5 bậc theo code ban đầu.
    # Thu nhập tính thuế <= 0 thì không phát sinh thuế.

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


def dinh_dang_tien(x):
    return f"{x:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")


def email_hop_le(email):
    if not email:
        return False
    return bool(
        re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(email).strip())
    )


def tao_du_lieu_ban_dau(so_luong):
    return pd.DataFrame({
        "Mã nhân viên": [
            f"NV{i:03d}" for i in range(1, so_luong + 1)
        ],
        "Họ và tên": ["" for _ in range(so_luong)],
        "Email": ["" for _ in range(so_luong)],
        "Lương cơ bản": [15.0 for _ in range(so_luong)],
        "Tiền thưởng": [0.0 for _ in range(so_luong)],
        "Phụ cấp tính BHXH": [0.0 for _ in range(so_luong)],
        "Phụ cấp không tính BHXH": [0.0 for _ in range(so_luong)],
        "Số người phụ thuộc": [0 for _ in range(so_luong)]
    })


# =========================================================
# 3. ĐỊNH DẠNG EXCEL
# =========================================================

def dinh_dang_sheet(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
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

    ws.row_dimensions[1].height = 30

    for col in ws.columns:
        do_rong = max(
            len(str(cell.value or "")) for cell in col
        )
        ws.column_dimensions[
            get_column_letter(col[0].column)
        ].width = min(do_rong + 3, 42)


def dinh_dang_cot_tien(ws, danh_sach_cot, cac_cot_tien):
    for ten_cot in cac_cot_tien:
        if ten_cot not in danh_sach_cot:
            continue

        vi_tri = danh_sach_cot.index(ten_cot) + 1

        for dong in range(2, ws.max_row + 1):
            ws.cell(
                row=dong,
                column=vi_tri
            ).number_format = '#,##0.000;[Red]-#,##0.000;0.000'


# =========================================================
# 4. TẠO BÁO CÁO LƯƠNG RIÊNG CHO TỪNG NHÂN VIÊN
# =========================================================

def tao_bao_cao_ca_nhan(nv):
    output = BytesIO()

    du_lieu = [
        ("Mã nhân viên", nv["Mã nhân viên"], ""),
        ("Họ và tên", nv["Họ và tên"], ""),
        ("Email", nv["Email"], ""),
        ("", "", ""),
        ("NỘI DUNG", "SỐ TIỀN", "ĐƠN VỊ"),
        ("Lương cơ bản", nv["Lương cơ bản"], "Triệu đồng"),
        ("Tiền thưởng", nv["Tiền thưởng"], "Triệu đồng"),
        ("Phụ cấp tính BHXH", nv["Phụ cấp tính BHXH"], "Triệu đồng"),
        (
            "Phụ cấp không tính BHXH",
            nv["Phụ cấp không tính BHXH"],
            "Triệu đồng"
        ),
        ("TỔNG THU NHẬP", nv["Tổng thu nhập"], "Triệu đồng"),
        ("Tiền đóng BHXH", nv["Tiền đóng BHXH"], "Triệu đồng"),
        ("Giảm trừ bản thân", nv["Giảm trừ bản thân"], "Triệu đồng"),
        ("Số người phụ thuộc", nv["Số người phụ thuộc"], "Người"),
        (
            "Giảm trừ người phụ thuộc",
            nv["Giảm trừ người phụ thuộc"],
            "Triệu đồng"
        ),
        (
            "THU NHẬP TÍNH THUẾ",
            nv["Thu nhập tính thuế"],
            "Triệu đồng"
        ),
        ("THUẾ TNCN", nv["Thuế TNCN"], "Triệu đồng"),
        ("TIỀN THỰC LÃNH", nv["Thực lãnh"], "Triệu đồng")
    ]

    df = pd.DataFrame(
        du_lieu,
        columns=["Nội dung", "Giá trị", "Đơn vị"]
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            sheet_name="Phieu luong",
            index=False
        )

        ws = writer.sheets["Phieu luong"]
        dinh_dang_sheet(ws)

        ws.column_dimensions["A"].width = 42
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 18

        for dong in [6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 17]:
            ws.cell(
                row=dong,
                column=2
            ).number_format = '#,##0.000;[Red]-#,##0.000;0.000'

        for dong in [10, 15, 16, 17]:
            for cell in ws[dong]:
                cell.font = Font(bold=True)

        dong_ghi_chu = ws.max_row + 2
        ws.cell(
            row=dong_ghi_chu,
            column=1,
            value=(
                "Đơn vị tiền: triệu đồng. "
                "Công thức tính theo mô hình được cài đặt trong ứng dụng."
            )
        )
        ws.cell(
            row=dong_ghi_chu + 1,
            column=1,
            value=(
                "Cần kiểm tra quy định thuế và bảo hiểm hiện hành "
                "trước khi sử dụng chính thức."
            )
        )

    output.seek(0)
    return output.getvalue()


# =========================================================
# 5. TẠO EXCEL TỔNG HỢP
# =========================================================

def tao_excel_tong_hop(bang):
    output = BytesIO()

    cac_cot = [
        "Mã nhân viên",
        "Họ và tên",
        "Email",
        "Lương cơ bản",
        "Tiền thưởng",
        "Phụ cấp tính BHXH",
        "Phụ cấp không tính BHXH",
        "Số người phụ thuộc",
        "Tổng thu nhập",
        "Tiền đóng BHXH",
        "Giảm trừ bản thân",
        "Giảm trừ người phụ thuộc",
        "Thu nhập tính thuế",
        "Thuế TNCN",
        "Thực lãnh"
    ]

    chi_tiet = bang[cac_cot].copy()

    tong_hop = pd.DataFrame({
        "Chỉ tiêu": [
            "Số người lao động",
            "Tổng thu nhập",
            "Tổng tiền đóng BHXH",
            "Tổng thu nhập tính thuế",
            "Tổng thuế TNCN",
            "Tổng tiền thực lãnh"
        ],
        "Giá trị": [
            len(bang),
            bang["Tổng thu nhập"].sum(),
            bang["Tiền đóng BHXH"].sum(),
            bang["Thu nhập tính thuế"].sum(),
            bang["Thuế TNCN"].sum(),
            bang["Thực lãnh"].sum()
        ],
        "Đơn vị": [
            "Người",
            "Triệu đồng",
            "Triệu đồng",
            "Triệu đồng",
            "Triệu đồng",
            "Triệu đồng"
        ]
    })

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        chi_tiet.to_excel(
            writer,
            sheet_name="Bang luong chi tiet",
            index=False
        )
        tong_hop.to_excel(
            writer,
            sheet_name="Tong hop quy luong",
            index=False
        )

        dinh_dang_sheet(writer.sheets["Bang luong chi tiet"])
        dinh_dang_sheet(writer.sheets["Tong hop quy luong"])

        dinh_dang_cot_tien(
            writer.sheets["Bang luong chi tiet"],
            cac_cot,
            [
                "Lương cơ bản",
                "Tiền thưởng",
                "Phụ cấp tính BHXH",
                "Phụ cấp không tính BHXH",
                "Tổng thu nhập",
                "Tiền đóng BHXH",
                "Giảm trừ bản thân",
                "Giảm trừ người phụ thuộc",
                "Thu nhập tính thuế",
                "Thuế TNCN",
                "Thực lãnh"
            ]
        )

        ws = writer.sheets["Tong hop quy luong"]
        for dong in range(3, ws.max_row + 1):
            ws.cell(
                row=dong,
                column=2
            ).number_format = '#,##0.000;[Red]-#,##0.000;0.000'

    output.seek(0)
    return output.getvalue()


# =========================================================
# 6. GỬI EMAIL KÈM PHIẾU LƯƠNG
# =========================================================

def gui_email_luong(
    smtp_server,
    smtp_port,
    email_cong_ty,
    mat_khau,
    email_nhan,
    ho_ten,
    ma_nv,
    file_excel
):
    msg = EmailMessage()
    msg["Subject"] = f"Phiếu lương cá nhân - {ma_nv}"
    msg["From"] = email_cong_ty
    msg["To"] = email_nhan

    msg.set_content(
        f"""Kính gửi {ho_ten or ma_nv},

Phòng nhân sự gửi bạn phiếu lương cá nhân đính kèm.

Vui lòng kiểm tra các khoản thu nhập, bảo hiểm, thuế TNCN
và số tiền thực lãnh. Nếu có sai sót, vui lòng liên hệ
phòng nhân sự.

Đề nghị bảo mật thông tin trong file đính kèm.

Trân trọng,
Phòng Nhân sự
"""
    )

    msg.add_attachment(
        file_excel,
        maintype="application",
        subtype=(
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        filename=f"Phieu_luong_{ma_nv}.xlsx"
    )

    if int(smtp_port) == 465:
        with smtplib.SMTP_SSL(
            smtp_server,
            int(smtp_port),
            context=ssl.create_default_context(),
            timeout=30
        ) as server:
            server.login(email_cong_ty, mat_khau)
            server.send_message(msg)
    else:
        with smtplib.SMTP(
            smtp_server,
            int(smtp_port),
            timeout=30
        ) as server:
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
            server.login(email_cong_ty, mat_khau)
            server.send_message(msg)


# =========================================================
# 7. CHỌN SỐ LƯỢNG VÀ NHẬP DỮ LIỆU
# =========================================================

st.header("1. Nhập thông tin nhân viên")

so_luong = int(st.number_input(
    "Số người lao động cần tính",
    min_value=1,
    max_value=500,
    value=10,
    step=1
))

key_du_lieu = f"du_lieu_luong_{so_luong}"

if key_du_lieu not in st.session_state:
    st.session_state[key_du_lieu] = tao_du_lieu_ban_dau(so_luong)

du_lieu_nhap = st.data_editor(
    st.session_state[key_du_lieu],
    key=f"editor_luong_{so_luong}",
    num_rows="fixed",
    use_container_width=True,
    hide_index=True,
    height=min(600, 38 * (so_luong + 1)),
    column_config={
        "Mã nhân viên": st.column_config.TextColumn(
            "Mã nhân viên",
            required=True
        ),
        "Họ và tên": st.column_config.TextColumn("Họ và tên"),
        "Email": st.column_config.TextColumn("Email nhân viên"),
        "Lương cơ bản": st.column_config.NumberColumn(
            "Lương cơ bản (triệu)",
            min_value=0.0,
            step=0.5,
            format="%.3f"
        ),
        "Tiền thưởng": st.column_config.NumberColumn(
            "Tiền thưởng (triệu)",
            min_value=0.0,
            step=0.5,
            format="%.3f"
        ),
        "Phụ cấp tính BHXH": st.column_config.NumberColumn(
            "Phụ cấp tính BHXH (triệu)",
            min_value=0.0,
            step=0.5,
            format="%.3f"
        ),
        "Phụ cấp không tính BHXH": st.column_config.NumberColumn(
            "Phụ cấp không tính BHXH (triệu)",
            min_value=0.0,
            step=0.5,
            format="%.3f"
        ),
        "Số người phụ thuộc": st.column_config.NumberColumn(
            "Người phụ thuộc",
            min_value=0,
            max_value=100,
            step=1,
            format="%d"
        )
    }
)


# =========================================================
# 8. TÍNH LƯƠNG CHO TOÀN BỘ NHÂN VIÊN
# =========================================================

if st.button(
    "🧮 TÍNH LƯƠNG",
    type="primary",
    use_container_width=True
):
    bang = du_lieu_nhap.copy()
    loi = []

    cot_tien = [
        "Lương cơ bản",
        "Tiền thưởng",
        "Phụ cấp tính BHXH",
        "Phụ cấp không tính BHXH"
    ]

    for cot in ["Mã nhân viên", "Họ và tên", "Email"]:
        bang[cot] = bang[cot].fillna("").astype(str).str.strip()

    if (bang["Mã nhân viên"] == "").any():
        loi.append("Mã nhân viên không được để trống.")

    if bang["Mã nhân viên"].duplicated().any():
        loi.append("Có mã nhân viên bị trùng.")

    for _, dong in bang.iterrows():
        if dong["Email"] and not email_hop_le(dong["Email"]):
            loi.append(
                f"Email không hợp lệ của nhân viên {dong['Mã nhân viên']}."
            )

    for cot in cot_tien:
        bang[cot] = pd.to_numeric(bang[cot], errors="coerce")

    bang["Số người phụ thuộc"] = pd.to_numeric(
        bang["Số người phụ thuộc"],
        errors="coerce"
    )

    if bang[cot_tien].isna().any().any():
        loi.append("Có khoản tiền trống hoặc không phải số.")

    if (bang[cot_tien] < 0).any().any():
        loi.append("Các khoản tiền không được âm.")

    if bang["Số người phụ thuộc"].isna().any():
        loi.append("Số người phụ thuộc không được để trống.")
    elif (
        (bang["Số người phụ thuộc"] < 0)
        | (bang["Số người phụ thuộc"] % 1 != 0)
    ).any():
        loi.append("Số người phụ thuộc phải là số nguyên không âm.")

    if loi:
        for thong_bao in loi:
            st.error(thong_bao)
    else:
        bang["Số người phụ thuộc"] = (
            bang["Số người phụ thuộc"].astype(int)
        )

        bang["Tổng thu nhập"] = (
            bang["Lương cơ bản"]
            + bang["Tiền thưởng"]
            + bang["Phụ cấp tính BHXH"]
            + bang["Phụ cấp không tính BHXH"]
        )

        bang["Tiền đóng BHXH"] = (
            bang["Lương cơ bản"]
            + bang["Phụ cấp tính BHXH"]
        ) * 0.105

        bang["Giảm trừ bản thân"] = 15.5

        bang["Giảm trừ người phụ thuộc"] = (
            bang["Số người phụ thuộc"] * 6.2
        )

        bang["Thu nhập tính thuế"] = (
            bang["Tổng thu nhập"]
            - bang["Tiền đóng BHXH"]
            - bang["Giảm trừ bản thân"]
            - bang["Giảm trừ người phụ thuộc"]
        )

        bang["Thuế TNCN"] = bang["Thu nhập tính thuế"].apply(
            tinh_thue_tncn
        )

        bang["Thực lãnh"] = (
            bang["Tổng thu nhập"]
            - bang["Tiền đóng BHXH"]
            - bang["Thuế TNCN"]
        )

        st.session_state["ket_qua_bang_luong"] = bang.copy()
        st.success(f"Đã tính lương cho {len(bang)} nhân viên.")


# =========================================================
# 9. HIỂN THỊ KẾT QUẢ VÀ XUẤT FILE
# =========================================================

if "ket_qua_bang_luong" in st.session_state:
    bang_kq = st.session_state["ket_qua_bang_luong"]

    st.divider()
    st.header("2. Bảng lương kết quả")

    cot_hien_thi = [
        "Mã nhân viên",
        "Họ và tên",
        "Email",
        "Tổng thu nhập",
        "Tiền đóng BHXH",
        "Giảm trừ bản thân",
        "Giảm trừ người phụ thuộc",
        "Thu nhập tính thuế",
        "Thuế TNCN",
        "Thực lãnh"
    ]

    st.dataframe(
        bang_kq[cot_hien_thi],
        use_container_width=True,
        hide_index=True
    )

    st.subheader("Tổng hợp quỹ lương")

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Số nhân viên", f"{len(bang_kq):,}")
    c2.metric(
        "Tổng thu nhập",
        f"{dinh_dang_tien(bang_kq['Tổng thu nhập'].sum())} triệu"
    )
    c3.metric(
        "Tổng thuế TNCN",
        f"{dinh_dang_tien(bang_kq['Thuế TNCN'].sum())} triệu"
    )
    c4.metric(
        "Tổng thực lãnh",
        f"{dinh_dang_tien(bang_kq['Thực lãnh'].sum())} triệu"
    )

    st.subheader("3. Tải báo cáo Excel")

    st.download_button(
        "📥 Tải Excel tổng hợp toàn bộ nhân viên",
        data=tao_excel_tong_hop(bang_kq),
        file_name="Bang_luong_tong_hop.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    zip_output = BytesIO()

    with zipfile.ZipFile(
        zip_output,
        "w",
        compression=zipfile.ZIP_DEFLATED
    ) as zip_file:
        ten_da_dung = set()

        for _, nv in bang_kq.iterrows():
            ma_nv = re.sub(
                r'[\\/*?:"<>|]',
                "_",
                str(nv["Mã nhân viên"]).strip()
            ).strip(" ._") or "Nhan_vien"

            ten_goc = ma_nv
            dem = 2

            while f"Phieu_luong_{ma_nv}.xlsx" in ten_da_dung:
                ma_nv = f"{ten_goc}_{dem}"
                dem += 1

            ten_file = f"Phieu_luong_{ma_nv}.xlsx"
            ten_da_dung.add(ten_file)

            zip_file.writestr(
                ten_file,
                tao_bao_cao_ca_nhan(nv)
            )

    zip_output.seek(0)

    st.download_button(
        "📦 Tải ZIP phiếu lương riêng của từng nhân viên",
        data=zip_output.getvalue(),
        file_name="Phieu_luong_tung_nhan_vien.zip",
        mime="application/zip",
        use_container_width=True
    )

    st.subheader("4. Xem và tải phiếu lương cá nhân")

    danh_sach_ma = bang_kq["Mã nhân viên"].astype(str).tolist()

    ma_chon = st.selectbox(
        "Chọn mã nhân viên",
        options=danh_sach_ma
    )

    nv_chon = bang_kq[
        bang_kq["Mã nhân viên"].astype(str) == ma_chon
    ].iloc[0]

    st.write(f"**Họ tên:** {nv_chon['Họ và tên'] or 'Chưa nhập'}")
    st.write(f"**Email:** {nv_chon['Email'] or 'Chưa nhập'}")

    st.download_button(
        f"📄 Tải phiếu lương của {ma_chon}",
        data=tao_bao_cao_ca_nhan(nv_chon),
        file_name=f"Phieu_luong_{ma_chon}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )


# =========================================================
# 10. CẤU HÌNH EMAIL VÀ GỬI PHIẾU LƯƠNG
# =========================================================

if "ket_qua_bang_luong" in st.session_state:
    bang_gui = st.session_state["ket_qua_bang_luong"]

    st.divider()
    st.header("5. Gửi phiếu lương qua email")

    st.info(
        "Nhập thông tin email công ty bên dưới. "
        "Mật khẩu chỉ tồn tại trong phiên nhập liệu; "
        "không ghi mật khẩu vào file Excel."
    )

    with st.form("form_gui_luong"):

        nha_cung_cap = st.selectbox(
            "Nhà cung cấp email",
            [
                "Gmail",
                "Microsoft 365 / Outlook",
                "Hostinger",
                "SMTP khác"
            ]
        )

        cau_hinh = {
            "Gmail": ("smtp.gmail.com", 587),
            "Microsoft 365 / Outlook": (
                "smtp.office365.com", 587
            ),
            "Hostinger": ("smtp.hostinger.com", 465),
            "SMTP khác": ("", 587)
        }

        smtp_mac_dinh, cong_mac_dinh = cau_hinh[nha_cung_cap]

        smtp_server = st.text_input(
            "Máy chủ SMTP",
            value=smtp_mac_dinh
        ).strip()

        smtp_port = st.number_input(
            "Cổng SMTP",
            min_value=1,
            max_value=65535,
            value=cong_mac_dinh,
            step=1
        )

        email_cong_ty = st.text_input(
            "Email công ty gửi bảng lương",
            placeholder="nhansu@congty.com"
        ).strip()

        mat_khau = st.text_input(
            "Mật khẩu ứng dụng / mật khẩu SMTP",
            type="password",
            help=(
                "Nên sử dụng App Password hoặc thông tin SMTP "
                "do quản trị viên cấp. Không sử dụng mật khẩu "
                "đăng nhập thông thường nếu nhà cung cấp không cho phép."
            )
        )

        danh_sach_nhan = bang_gui[
            bang_gui["Email"].fillna("").astype(str).str.strip() != ""
        ].copy()

        danh_sach_nhan = danh_sach_nhan[
            danh_sach_nhan["Email"].apply(email_hop_le)
        ]

        pham_vi_gui = st.radio(
            "Chọn người nhận",
            [
                "Gửi cho một nhân viên",
                "Gửi cho tất cả nhân viên có email hợp lệ"
            ]
        )

        ma_nv_chon = None

        if pham_vi_gui == "Gửi cho một nhân viên":
            if not danh_sach_nhan.empty:
                lua_chon = danh_sach_nhan[
                    ["Mã nhân viên", "Họ và tên", "Email"]
                ].to_dict("records")

                nv_gui_chon = st.selectbox(
                    "Chọn nhân viên nhận phiếu lương",
                    options=lua_chon,
                    format_func=lambda nv: (
                        f"{nv['Mã nhân viên']} - "
                        f"{nv['Họ và tên']} ({nv['Email']})"
                    )
                )

                ma_nv_chon = nv_gui_chon["Mã nhân viên"]
            else:
                st.warning("Chưa có email nhân viên hợp lệ.")

        xac_nhan_gui = st.checkbox(
            "Tôi đã kiểm tra danh sách người nhận và xác nhận "
            "gửi đúng phiếu lương cho từng nhân viên."
        )

        bam_gui = st.form_submit_button(
            "📨 GỬI PHIẾU LƯƠNG",
            type="primary",
            use_container_width=True
        )

    if bam_gui:

        if not smtp_server or not email_cong_ty or not mat_khau:
            st.error("Vui lòng nhập đủ SMTP, email và mật khẩu ứng dụng.")

        elif not email_hop_le(email_cong_ty):
            st.error("Email công ty không đúng định dạng.")

        elif not xac_nhan_gui:
            st.error("Bạn phải xác nhận danh sách người nhận.")

        elif danh_sach_nhan.empty:
            st.error("Không có nhân viên nào có email hợp lệ.")

        else:
            if pham_vi_gui == "Gửi cho một nhân viên":
                danh_sach_can_gui = danh_sach_nhan[
                    danh_sach_nhan["Mã nhân viên"] == ma_nv_chon
                ]
            else:
                danh_sach_can_gui = danh_sach_nhan

            ket_qua_gui = []
            tien_trinh = st.progress(0)
            thong_bao = st.empty()

            for i, (_, nv) in enumerate(
                danh_sach_can_gui.iterrows(),
                start=1
            ):
                ma_nv = str(nv["Mã nhân viên"]).strip()
                email_nhan = str(nv["Email"]).strip()
                ho_ten = str(nv["Họ và tên"]).strip()

                try:
                    file_ca_nhan = tao_bao_cao_ca_nhan(nv)

                    gui_email_luong(
                        smtp_server=smtp_server,
                        smtp_port=int(smtp_port),
                        email_cong_ty=email_cong_ty,
                        mat_khau=mat_khau,
                        email_nhan=email_nhan,
                        ho_ten=ho_ten,
                        ma_nv=ma_nv,
                        file_excel=file_ca_nhan
                    )

                    ket_qua_gui.append({
                        "Mã nhân viên": ma_nv,
                        "Họ và tên": ho_ten,
                        "Email": email_nhan,
                        "Kết quả": "Thành công",
                        "Chi tiết": "Máy chủ SMTP đã chấp nhận email"
                    })

                except Exception as e:
                    ket_qua_gui.append({
                        "Mã nhân viên": ma_nv,
                        "Họ và tên": ho_ten,
                        "Email": email_nhan,
                        "Kết quả": "Thất bại",
                        "Chi tiết": str(e)[:300]
                    })

                tien_trinh.progress(
                    i / len(danh_sach_can_gui)
                )
                thong_bao.write(
                    f"Đã xử lý {i}/{len(danh_sach_can_gui)} email."
                )

            df_gui = pd.DataFrame(ket_qua_gui)
            st.session_state["ket_qua_gui_email"] = df_gui

            thanh_cong = (
                df_gui["Kết quả"] == "Thành công"
            ).sum()

            st.success(
                f"Hoàn tất: {thanh_cong} thành công, "
                f"{len(df_gui) - thanh_cong} thất bại."
            )

    if "ket_qua_gui_email" in st.session_state:
        st.subheader("Kết quả gửi email")

        st.dataframe(
            st.session_state["ket_qua_gui_email"],
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "📥 Tải báo cáo trạng thái gửi email",
            data=st.session_state["ket_qua_gui_email"].to_csv(
                index=False
            ).encode("utf-8-sig"),
            file_name="Ket_qua_gui_email.csv",
            mime="text/csv"
        )


st.caption(
    "Bảo mật: Phiếu lương chứa dữ liệu cá nhân. "
    "Chỉ gửi đúng phiếu lương cho người nhận tương ứng."
)
