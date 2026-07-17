# Auto-Wash — Bộ tài liệu yêu cầu phần mềm (SWR302)

Hệ thống quản lý trung tâm rửa xe tự động: khách hàng đặt lịch qua Mobile Web, nhân viên vận hành hàng chờ, nhận dạng biển số qua Camera LPR + 3rd-party LPR API, thông báo qua Email SMTP. Phạm vi gồm 5 service nội bộ: Account, Booking, Queue, Notification, LPR.

Toàn bộ tài liệu trong folder này **đã đồng bộ với nhau** tính đến 16/07/2026 — bản chính thức: **SRS v2.6**.

## Cấu trúc folder

```
Auto-Wash-Project/
├── README.md                          (file này)
├── Auto-Wash_Plan_CaiTien.xlsx        Kế hoạch cải tiến + ma trận truy vết + checklist đồng bộ
├── docs/
│   └── Auto-Wash_SRS_v2.6.docx        SRS chính thức, tiếng Việt (bản dùng)
├── diagrams/
│   ├── ContextDiagram_v2.drawio       Sơ đồ ngữ cảnh
│   ├── UseCaseDiagram.drawio          Use case diagram (5 page, đã gắn mã UC1–UC24)
│   ├── ActivityDiagrams.drawio        Activity diagram (5 page, ngang)
│   ├── Swimlanes.drawio               Swimlane diagram (5 page, ngang)
│   ├── StateTransition.drawio         Sơ đồ chuyển trạng thái (1 page)
│   ├── ERD.drawio                     Sơ đồ thực thể — 9 thực thể
│   ├── ERD.png                        Ảnh ERD render sẵn (chèn vào SRS mục 2.2)
│   ├── Mockups.html                   Gallery xem nhanh 23 mockup
│   ├── DialogMap.drawio               Dialog map — 2 sheet: Khách hàng, Nhân viên
│   ├── DialogMap_Customer.png         Ảnh sheet Khách hàng (chèn SRS 4.1.1)
│   ├── DialogMap_Staff.png            Ảnh sheet Nhân viên (chèn SRS 4.1.2)
│   └── mockups/                       PNG 23 màn mockup (nguồn chèn vào SRS mục 4)
├── legacy/                            Bản cũ, chỉ để tham chiếu
│   ├── ContextDiagram_v1.drawio
│   ├── Auto-Wash_SRS_v1.0_Full.docx   (tiếng Anh, còn loyalty/tier)
│   ├── Auto-Wash_SRS_v2.0.docx
│   ├── Auto-Wash_SRS_v2.1.docx
│   ├── Auto-Wash_SRS_v2.2.docx
│   ├── Auto-Wash_SRS_v2.3.docx
│   ├── Auto-Wash_SRS_v2.4.docx
│   └── Auto-Wash_SRS_v2.5.docx
└── tools/srs-generator/               Script sinh file SRS docx
    ├── build.js  helpers.js  uc1.js  uc2.js
```

## Mô tả từng file

### docs/Auto-Wash_SRS_v2.6.docx — tài liệu trung tâm

SRS đầy đủ ~54 trang tiếng Việt, cấu trúc theo mẫu FA / IEEE 830:

1. **Giới thiệu** — mục đích, phạm vi, đối tượng đọc, từ viết tắt, 1.4.1 Bảng thuật ngữ nghiệp vụ (Glossary), tài liệu tham khảo.
2. **Yêu cầu tổng quát** — sơ đồ ngữ cảnh + bảng luồng tương tác, ERD 9 thực thể (đã chèn hình), workflow, 2 bảng chuyển trạng thái, bảng use case kèm ghi chú «include»/«extend», ma trận phân quyền.
3. **Đặc tả 24 Use Case** — mỗi UC gồm: Mục tiêu, Tác nhân, Kích hoạt, Điều kiện tiên quyết, Kết quả, **Luồng sự kiện chính**, **Luồng thay thế / ngoại lệ**, **Quy tắc nghiệp vụ** (BR đánh số toàn cục).
4. **Màn hình Mockup** — 4.1 Sơ đồ luồng màn hình (Dialog Map: Khách hàng + Nhân viên); 4.2–4.6 cả 23 màn wireframe (xem nhanh tại `diagrams/Mockups.html`).
5. **Yêu cầu phi chức năng** — hiệu năng, bảo mật, khả dụng, sẵn sàng/tin cậy, mở rộng.
6. **Tích hợp** — Supabase PostgreSQL, Gmail SMTP/MailKit, LPR API, React Mobile Web.
7. **Di trú dữ liệu**.
8. **Phụ lục** — 32 thông điệp MSG, 3 mẫu email ET.

**Điểm nghiệp vụ then chốt:** đăng nhập **chỉ dùng mật khẩu** (OTP chỉ còn cho đăng ký, đổi mật khẩu, đăng ký xe); đặt lịch có bước **xem tóm tắt trước khi xác nhận**; đăng ký báo email trùng thì **điều hướng sang đăng nhập**; hàng chờ **FIFO**; giá = tổng giá dịch vụ (không giảm giá). So với v1.0 đã **loại bỏ** loyalty/tier/điểm, voucher, campaign, admin dashboard, Google OAuth.

**Lịch sử phiên bản (tóm tắt):**

- **v2.0** — dịch/viết lại từ v1.0 sang tiếng Việt đúng phạm vi ContextDiagram.
- **v2.1** — đồng bộ theo Use Case Diagram (quan hệ «include»/«extend», điều hướng, bước tóm tắt đặt lịch).
- **v2.2** — bổ sung 7 nhánh lỗi/ngoại lệ và MSG26–MSG32 cho 24 UC.
- **v2.3** — đồng bộ toàn bộ sơ đồ: mã UC1–UC24, nhánh UC19, luồng notifications trên Context, ghi chú State, Glossary; vẽ ERD; tạo 23 mockup.
- **v2.4** — bỏ OTP khi đăng nhập (UC1 chỉ dùng mật khẩu); đồng bộ UseCase/Activity/Swimlane/State/mockup/phụ lục.
- **v2.5** — hoàn thiện đặc tả UC: mỗi UC có Luồng sự kiện chính + Luồng thay thế/ngoại lệ + Quy tắc nghiệp vụ.
- **v2.6** — bổ sung mục 4.1 Dialog Map (luồng màn hình Khách hàng & Nhân viên); danh mục mockup dời xuống 4.2–4.6.

### diagrams/ContextDiagram_v2.drawio

Sơ đồ ngữ cảnh: Auto-Wash System (5 service) ở giữa; tác nhân Staff, Customer; kênh Mobile Web; hệ thống ngoài Email SMTP, Camera LPR, 3rd-party LPR API. Có đủ luồng request plate scan, các luồng phản hồi cho Customer, update queue status, view vehicles & services, manage account và view & mark notifications.

### diagrams/UseCaseDiagram.drawio — 5 page theo phân hệ

24 UC (đã gắn **mã UC1–UC24**) + sub-UC "Verify OTP via email". «Track queue progress» là nhánh con của UC13 (không tính là UC độc lập → tổng đúng 24). Quan hệ: **3 «include»** (Register, Change password, Register vehicle → Verify OTP — Login **không** dùng OTP); **4 «extend»** (Delete vehicle → View vehicles; Track queue progress → View bookings; Scan LPR → Check-in; Mark as read → View notifications). Nguyên tắc: bước bắt buộc actor cảm nhận được → include, hành vi tùy chọn/có điều kiện → extend; không đưa bước xử lý nội bộ lên diagram.

### diagrams/ActivityDiagrams.drawio và Swimlanes.drawio — mỗi file 5 page

Cùng 17 luồng, nhóm theo cụm: (1) Account: Register + Login + Update profile + Change password; (2) Vehicle: view/register/delete; (3) Booking: create + view/track + confirm + cancel; (4) Queue: check-in→check-out + cancel + update; (5) Administration + Notifications. ActivityDiagrams vẽ trục ngang đơn; Swimlanes vẽ làn ngang (Customer / Staff / Auto-Wash System / Email SMTP).

Quy ước: chuẩn UML trắng đen (chấm đen start, bullseye end, thoi decision); vòng lặp lỗi vòng lên trên, nhánh từ chối đi xuống dưới; nhãn không dùng mã MSG. Tình huống biên đã phủ: OTP hết hạn ≠ OTP sai, tài khoản bị khóa, email trùng → nhảy sang Login, chưa có xe khi đặt lịch, re-check Pending khi confirm (concurrency), chuyển trạng thái hàng chờ không hợp lệ (UC19), LPR quá 5 giây → nhập tay, walk-in nhập tiền tay, không tài khoản → không gửi email, chặn tự khóa tài khoản mình. Luồng "B. Login" chỉ còn định danh + mật khẩu → tạo session (đã bỏ OTP).

### diagrams/StateTransition.drawio — 1 page, 3 sơ đồ

Booking (Pending → Confirmed → CheckedIn → Completed, nhánh Cancelled — khớp bảng 2.4.1 SRS); Queue (composite "In progress" bao 5 công đoạn, một mũi tên Cancelled với guard [not Completed], kèm ghi chú ràng buộc thứ tự công đoạn UC19 — khớp 2.4.2); OTP lifecycle (Active → Verified/Expired, nguồn UC3/UC5/UC8 — khớp CBR3).

### diagrams/ERD.drawio — 9 thực thể

Account, Customer, Vehicle, Booking, BookingService, Service, Queue, Notification, OtpVerification (kèm trường chính và quan hệ). `ERD.png` là bản render để chèn vào SRS mục 2.2; muốn hình chuẩn drawio thì mở `ERD.drawio` và File > Export as PNG đè lên `ERD.png` rồi build lại.

### diagrams/Mockups.html + mockups/ — 23 màn wireframe

14 màn khách hàng (mobile): landing, login, register(_otp), profile, change_password, my_vehicles, add_vehicle_otp, services_list, create_booking, booking_summary, my_bookings, notifications(_detail). 9 màn nhân viên (desktop): manage_users, manage_services, pending_bookings, staff_queue, advance_status, update_queue, checkin/checkout/cancel dialog. Mỗi màn kèm ghi chú nghiệp vụ (UC/MSG tương ứng). `Mockups.html` để xem nhanh cả bộ; PNG được `build.js` chèn vào SRS mục 4.

### tools/srs-generator/ — sinh lại file SRS docx

`build.js` chứa nội dung các mục 1, 2, 4–8 + lắp ráp document; `uc1.js` / `uc2.js` chứa dữ liệu 24 đặc tả UC (BR tự đánh số khi render); `helpers.js` là hàm dựng bảng/heading và `renderUC` (Luồng sự kiện chính + Luồng thay thế/ngoại lệ + Quy tắc nghiệp vụ). Sinh lại file:

```
npm install docx
node build.js        # tạo Auto-Wash_SRS_v2.6.docx tại chỗ (tự chèn ERD.png + DialogMap_*.png + mockups/*.png nếu có)
```

**Quan trọng:** muốn sửa nội dung SRS thì sửa các file `.js` rồi build lại, **đừng sửa tay file .docx** để tránh lệch nguồn.

## Nhật ký công việc đã làm

1. Viết lại SRS v1.0 (tiếng Anh) thành **v2.0 tiếng Việt** đúng phạm vi ContextDiagram: bỏ loyalty/tier, voucher, campaign, dashboard, Google OAuth; thêm LPR, Quản lý dịch vụ/tài khoản, Xác nhận/Hủy đặt lịch cho Staff.
2. Review & sửa sơ đồ ngữ cảnh thành **ContextDiagram_v2** (luồng phản hồi Customer, đổi tên 3rd-party LPR API, request plate scan, các luồng gộp phủ đủ UC).
3. Vẽ **UseCaseDiagram** (tách 5 page, chuẩn hóa include/extend, tên tiếng Anh).
4. Đồng bộ SRS lên **v2.1** (quan hệ 2.5, điều hướng UC3, bước tóm tắt UC12, bảng 2.1 khớp context).
5. Vẽ **ActivityDiagrams + Swimlanes** (17 luồng: bỏ màu, chuyển ngang, gộp cụm, chống chồng chéo).
6. Vẽ **StateTransition** (Booking, Queue, OTP).
7. Dọn dẹp: xóa BusinessFlow.drawio (bản luồng cũ).
8. Rà soát SRS lên **v2.2**: bổ sung nhánh lỗi/ngoại lệ + 7 BR/MSG (MSG26–32).
9. **Đồng bộ & hoàn thiện lên v2.3** (theo `Auto-Wash_Plan_CaiTien.xlsx`): mã UC1–UC24, nhánh UC19 trên Activity/Swimlane/State, luồng notifications trên Context, Glossary; vẽ ERD; tạo 23 mockup; chèn ERD + mockup vào SRS.
10. **Bỏ OTP khi đăng nhập → v2.4:** UC1 chỉ dùng định danh + mật khẩu; đồng bộ UseCaseDiagram, Activity/Swimlane (rút gọn luồng Login), StateTransition, mockup Đăng nhập, ET1 và MSG4.
11. **Hoàn thiện đặc tả UC → v2.5:** nâng `renderUC` để mỗi UC có Luồng sự kiện chính + Luồng thay thế/ngoại lệ + Quy tắc nghiệp vụ; build lại (đủ 24 UC spec, 24 hình).
12. **Bổ sung Dialog Map → v2.6:** vẽ 2 sơ đồ luồng màn hình (Khách hàng Mobile Web, Nhân viên Desktop) gộp trong 1 file DialogMap.drawio (2 sheet) + 2 PNG; chèn vào SRS mục 4.1, dời danh mục mockup xuống 4.2–4.6; build lại (26 hình).

## Việc còn dở / gợi ý bước tiếp theo

- Các sơ đồ Context/UseCase/Activity/Swimlane/State trong SRS vẫn ghi [xem file drawio]: khi cần, export PNG từ drawio rồi map vào `build.js` giống ERD/mockup.
- Khi đổi nghiệp vụ, sửa theo thứ tự: **ContextDiagram → UseCaseDiagram → Activity/Swimlane/StateTransition → SRS (qua srs-generator)** để giữ đồng bộ; luôn build lại và rà soát sau mỗi lần sửa.
