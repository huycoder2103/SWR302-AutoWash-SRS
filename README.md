<div align="center">

# 🚗 Auto-Wash

### Hệ thống quản lý trung tâm rửa xe tự động
**Software Requirements Specification (SRS) — SWR302**

![SRS](https://img.shields.io/badge/SRS-v2.9-1f3864)
![Use%20Cases](https://img.shields.io/badge/Use%20Cases-24-2563eb)
![Diagrams](https://img.shields.io/badge/Diagrams-9-7c3aed)
![Mockups](https://img.shields.io/badge/Mockups-23-059669)
![Pages](https://img.shields.io/badge/SRS%20Pages-65-b45309)
![Status](https://img.shields.io/badge/status-completed-success)

[**🌐 Xem trang tổng hợp (index.html)**](./index.html) &nbsp;·&nbsp; [**📄 Tài liệu SRS (PDF)**](./docs/Auto-Wash_SRS_v2.9.pdf) &nbsp;·&nbsp; [**🖥️ Slide thuyết trình**](./thuyet-trinh/Slide_ThuyetTrinh_AutoWash.html)

</div>

---

## 📌 Giới thiệu

**Auto-Wash** là hệ thống quản lý trung tâm rửa xe tự động: khách hàng đặt lịch qua **Mobile Web**, nhân viên vận hành **hàng chờ** rửa xe, nhận dạng biển số bằng **Camera LPR + 3rd-party LPR API**, và gửi thông báo qua **Email SMTP**.

Repository này chứa **toàn bộ tài liệu đặc tả yêu cầu phần mềm (SRS)** cùng bộ sơ đồ thiết kế, mockup, slide thuyết trình và trang web tổng hợp để deploy GitHub Pages.

> Phạm vi gồm 5 service nội bộ: **Account · Booking · Queue · Notification · LPR**.

## ✨ Phạm vi chức năng

| Nhóm | Chức năng chính |
|---|---|
| **Tài khoản** | Đăng ký (OTP), đăng nhập (chỉ mật khẩu), đổi mật khẩu, quản lý hồ sơ & tài khoản |
| **Phương tiện** | Đăng ký xe (OTP), xem/xóa xe |
| **Dịch vụ & Đặt lịch** | Xem dịch vụ, tạo/xác nhận/hủy đặt lịch, xem tóm tắt trước khi xác nhận |
| **Hàng chờ** | Check-in (LPR), chuyển trạng thái theo công đoạn, check-out, hủy lượt (FIFO) |
| **Thông báo** | Gửi & xem thông báo, email OTP/xác nhận/hoàn tất |

**Đã loại bỏ** (ngoài phạm vi): loyalty/tích điểm, voucher, campaign, admin dashboard, Google OAuth.

## 🧭 Sơ đồ thiết kế (tương tác)

Mở trực tiếp trên trình duyệt qua trang chủ, hoặc các link dưới đây:

| # | Sơ đồ | Link |
|---|---|---|
| 1 | Sơ đồ ngữ cảnh (Context) | [context.html](./diagrams-html/context.html) |
| 2 | Use Case (24 UC / 5 phân hệ) | [usecase.html](./diagrams-html/usecase.html) |
| 3 | Thực thể (ERD — 9 thực thể) | [erd.html](./diagrams-html/erd.html) |
| 4 | Chuyển trạng thái (Booking/Queue/OTP) | [state.html](./diagrams-html/state.html) |
| 5 | Hoạt động (Activity — 5 phân hệ) | [activity.html](./diagrams-html/activity.html) |
| 6 | Phân làn (Swimlane — 5 phân hệ) | [swimlane.html](./diagrams-html/swimlane.html) |
| 7 | Luồng dữ liệu (DFD Level 1) | [dataflow.html](./diagrams-html/dataflow.html) |
| 8 | Dialog Map (Khách hàng & Nhân viên) | [dialogmap.html](./diagrams-html/dialogmap.html) |
| 9 | Mockup giao diện (23 màn) | [mockups.html](./diagrams-html/mockups.html) |

## 🛠️ Công nghệ (theo thiết kế)

| Thành phần | Công nghệ |
|---|---|
| Frontend | React 18 + Vite (Mobile Web) |
| Backend | ASP.NET Core 8, Entity Framework Core |
| Cơ sở dữ liệu | PostgreSQL (Supabase) |
| Email | Gmail SMTP + MailKit |
| Nhận dạng biển số | Camera LPR + 3rd-party LPR API |
| Sơ đồ | draw.io (diagrams.net) |
| Sinh tài liệu | Node.js + `docx` (srs-generator) |

## 📂 Cấu trúc repository

```
Auto-Wash-Project/
├── index.html                     # Trang tổng hợp (GitHub Pages)
├── docs/
│   ├── Auto-Wash_SRS_v2.9.pdf      # SRS bản PDF (mở trực tiếp trên browser)
│   └── Auto-Wash_SRS_v2.9.docx     # SRS bản Word (nguồn chỉnh sửa)
├── diagrams-html/                 # 9 trang sơ đồ tương tác (deploy) + mockups/
├── diagrams/
│   ├── *.drawio                    # File sơ đồ gốc (draw.io)
│   └── diagram picture/            # Ảnh export từ draw.io (PNG/HTML)
├── thuyet-trinh/
│   ├── Slide_ThuyetTrinh_AutoWash.html   # Slide 25 trang (kiểu PowerPoint)
│   ├── CheatSheet_AutoWash.docx          # Cheat sheet + phân công 4 người
│   └── Script_ThuyetTrinh_AutoWash.docx  # Script thuyết trình
├── tools/srs-generator/           # Sinh lại SRS (build.js, helpers.js, uc1/uc2.js)
├── Auto-Wash_Plan_CaiTien.xlsx    # Kế hoạch cải tiến + ma trận truy vết
└── legacy/                        # Bản cũ (tham chiếu)
```

## 📄 Nội dung tài liệu SRS

Tài liệu **~65 trang**, cấu trúc theo mẫu FA / IEEE 830:

1. **Giới thiệu** — mục đích, phạm vi, thuật ngữ, Glossary, **giả định & ràng buộc**
2. **Yêu cầu tổng quát** — Context, ERD + **Từ điển dữ liệu**, Activity/Swimlane/DFD, State, Use Case, ma trận phân quyền
3. **Đặc tả 24 Use Case** — luồng chính, luồng thay thế/ngoại lệ, quy tắc nghiệp vụ
4. **Mockup** — Dialog Map + 23 màn giao diện
5. **Yêu cầu phi chức năng** — hiệu năng, bảo mật, khả dụng, bảo trì, khả chuyển, tuân thủ, **giao diện ngoài**
6. **Tích hợp** · 7. **Di trú dữ liệu** · 8. **Phụ lục** (32 MSG, 3 mẫu email)

## 🚀 Deploy GitHub Pages

1. Commit & push toàn bộ repository lên GitHub.
2. Vào **Settings → Pages** → Source: **Deploy from a branch** → Branch `main`, folder **`/ (root)`** → **Save**.
3. Truy cập: `https://<username>.github.io/<repo-name>/`

## 🔧 Sinh lại tài liệu SRS

```bash
cd tools/srs-generator
npm install docx
node build.js        # tạo Auto-Wash_SRS_v2.9.docx (tự chèn ảnh sơ đồ + mockup)
```

> Sửa nội dung SRS ở `uc1.js` / `uc2.js` / `build.js` rồi build lại — **không sửa tay file .docx** để tránh lệch nguồn.

## 👥 Nhóm thực hiện

| Thành viên | MSSV | Vai trò |
|---|---|---|
| Nguyễn Hoàng Huy | SE190240 | Trưởng nhóm · SRS · sơ đồ |
| Nguyễn Lê Thuận | SE190305 | Use Case · đặc tả |
| Võ Lê Trung Nguyên | SE190220 | ERD · dữ liệu · trạng thái |
| Nguyễn Thành Đạt | SE190239 | Mockup · giao diện · trình bày |

<div align="center">

**FPT University · SWR302 · TP. Hồ Chí Minh, 07/2026**

</div>
