# Auto-Wash — Bộ tài liệu yêu cầu phần mềm (SWP391)

Hệ thống quản lý trung tâm rửa xe tự động: khách hàng đặt lịch qua Mobile Web, nhân viên vận hành hàng chờ, nhận dạng biển số qua Camera LPR + 3rd-party LPR API, thông báo qua Email SMTP. Phạm vi gồm 5 service nội bộ: Account, Booking, Queue, Notification, LPR.

Toàn bộ tài liệu trong folder này **đã đồng bộ với nhau** tính đến 16/07/2026 (bản v2.4 — đăng nhập đã bỏ OTP).

## Cấu trúc folder

```
Auto-Wash-Project/
├── README.md                  (file này)
├── Auto-Wash_Plan_CaiTien.xlsx        Kế hoạch cải tiến + ma trận truy vết + checklist đồng bộ
├── docs/
│   └── Auto-Wash_SRS_v2.4.docx        SRS chính thức, tiếng Việt (bản dùng — đăng nhập bỏ OTP, đủ ERD + 23 mockup)
├── diagrams/
│   ├── ContextDiagram_v2.drawio       Sơ đồ ngữ cảnh (đã thêm luồng view & mark notifications)
│   ├── UseCaseDiagram.drawio          Use case diagram (5 page, đã gắn mã UC1–UC24)
│   ├── ActivityDiagrams.drawio        Activity diagram (5 page, ngang; UC19 đã thêm nhánh transition)
│   ├── Swimlanes.drawio               Swimlane diagram (5 page, ngang; UC19 đã thêm nhánh transition)
│   ├── StateTransition.drawio         Sơ đồ chuyển trạng thái (1 page; ghi chú ràng buộc UC19)
│   ├── ERD.drawio                     Sơ đồ thực thể — 9 thực thể (mới)
│   ├── ERD.png                        Ảnh ERD render sẵn để chèn SRS (mới)
│   ├── Mockups.html                   Gallery wireframe các màn hình chính (mới)
│   └── mockups/                       PNG cả 23 màn mockup (SVG render): login, landing, register(_otp), profile, change_password, my_vehicles, add_vehicle_otp, services_list, create_booking, booking_summary, my_bookings, notifications(_detail); manage_users, manage_services, pending_bookings, staff_queue, advance_status, update_queue, checkin/checkout/cancel dialog
├── legacy/                            Bản cũ, chỉ để tham chiếu
│   ├── ContextDiagram_v1.drawio
│   ├── Auto-Wash_SRS_v1.0_Full.docx   (tiếng Anh, còn loyalty/tier)
│   ├── Auto-Wash_SRS_v2.0.docx
│   ├── Auto-Wash_SRS_v2.1.docx
│   ├── Auto-Wash_SRS_v2.2.docx
│   └── Auto-Wash_SRS_v2.3.docx
└── tools/srs-generator/               Script sinh file SRS docx
    ├── build.js  helpers.js  uc1.js  uc2.js
```

## Mô tả từng file

### docs/Auto-Wash_SRS_v2.4.docx — tài liệu trung tâm

SRS đầy đủ ~45 trang tiếng Việt, cấu trúc theo mẫu FA/IEEE 830: (1) Giới thiệu (kèm 1.4.1 Glossary chuẩn hóa); (2) Yêu cầu tổng quát — sơ đồ ngữ cảnh + bảng luồng tương tác, **ERD 9 thực thể (đã chèn hình)**, workflow, 2 bảng chuyển trạng thái, bảng use case kèm ghi chú «include»/«extend», ma trận phân quyền; (3) Đặc tả 24 use case với business rule đánh số toàn cục; (4) Danh mục mockup — **cả 23 màn hình đã chèn wireframe (SVG)**, xem nhanh tại Mockups.html; (5) Yêu cầu phi chức năng; (6) Tích hợp; (7) Di trú dữ liệu; (8) Phụ lục — 32 thông điệp MSG, 3 mẫu email ET.

**v2.4 (16/07/2026) — bỏ OTP khi đăng nhập:** UC1 chỉ dùng định danh + mật khẩu; đồng bộ UseCaseDiagram (bỏ include Login→OTP), Activity/Swimlane (rút gọn luồng Login), StateTransition (OTP nguồn UC3/UC5/UC8), mockup Đăng nhập, ET1 và MSG4. OTP vẫn dùng cho UC3/UC5/UC8.

**v2.3 (15/07/2026) — đồng bộ sơ đồ với SRS:** (1) UseCaseDiagram gắn mã UC1–UC24 + chú thích «Track queue progress» là nhánh con của UC13 (chốt đúng 24 UC); (2) Activity & Swimlane bổ sung nhánh kiểm tra chuyển trạng thái hàng chờ hợp lệ cho UC19 (khớp MSG31); (3) ContextDiagram thêm luồng view & mark notifications (UC23/UC24); (4) StateTransition thêm ghi chú ràng buộc thứ tự công đoạn (UC19); (5) bổ sung 1.4.1 Glossary + ghi chú các UC chủ đích không vẽ Activity (UC2, UC10, UC16); (6) vẽ ERD 9 thực thể và chèn vào mục 2.2; (7) tạo 6 mockup wireframe chèn vào mục 4.

**v2.2 (12/07/2026):** rà soát toàn bộ 24 UC, bổ sung 7 MSG mới (MSG26–32) cho các nhánh lỗi/ngoại lệ: UC1 chặn đăng nhập khi tài khoản bị khóa; UC6 chặn tự vô hiệu hóa tài khoản mình; UC12 chặn đặt lịch khi chưa có xe; UC14 concurrency khi xác nhận booking; UC15 báo lỗi khi hủy booking sai trạng thái; UC19 báo lỗi khi chuyển trạng thái hàng chờ không hợp lệ; UC11 báo lỗi khi giá/thời lượng dịch vụ không hợp lệ.

Điểm nghiệp vụ then chốt: đăng nhập **chỉ dùng mật khẩu** (đã bỏ OTP từ v2.4; OTP chỉ còn cho đăng ký, đổi mật khẩu, đăng ký xe); đặt lịch có bước **xem tóm tắt trước khi xác nhận**; đăng ký báo email trùng thì **điều hướng sang đăng nhập**; hàng chờ **FIFO**; giá = tổng giá dịch vụ (không giảm giá); phiên bản này **đã loại bỏ** loyalty/tier/điểm, voucher, campaign, admin dashboard, Google OAuth so với v1.0.

### diagrams/ContextDiagram_v2.drawio

Sơ đồ ngữ cảnh: Auto-Wash System (5 service) ở giữa; tác nhân Staff, Customer; kênh Mobile Web; hệ thống ngoài Email SMTP, Camera LPR, 3rd-party LPR API. So với v1: đổi tên "License Plate Service" → "3rd-party LPR API", thêm luồng request plate scan (System → Camera), 3 luồng phản hồi cho Customer, luồng update queue status / view vehicles & services / manage account.

### diagrams/UseCaseDiagram.drawio — 5 page theo phân hệ

24 UC (đã gắn **mã UC1–UC24**) + sub-UC "Verify OTP via email", tên tiếng Anh. «Track queue progress» được chú thích là nhánh con của UC13 (không tính là UC độc lập → tổng đúng 24). Quan hệ: 3 «include» (Register, Change password, Register vehicle → Verify OTP — **Login đã bỏ include OTP từ v2.4**); 4 «extend» (Delete vehicle → View vehicles; Track queue progress → View bookings; Scan LPR → Check-in; Mark as read → View notifications). Nguyên tắc: bước bắt buộc actor cảm nhận được → include, hành vi tùy chọn/có điều kiện → extend; không đưa bước xử lý nội bộ lên diagram.

### diagrams/ActivityDiagrams.drawio và Swimlanes.drawio — mỗi file 5 page

Cùng 17 luồng, nhóm theo cụm: (1) Account: Register + Login + Update profile + Change password; (2) Vehicle: view/register/delete; (3) Booking: create + view/track + confirm + cancel; (4) Queue: check-in→check-out + cancel + update; (5) Administration + Notifications. ActivityDiagrams vẽ dạng trục ngang đơn; Swimlanes vẽ làn ngang (Customer / Staff / Auto-Wash System / Email SMTP), mỗi pool chỉ chứa làn segment đó dùng.

Quy ước: chuẩn UML trắng đen (chấm đen start, bullseye end, thoi decision); vòng lặp lỗi vòng lên trên, nhánh từ chối đi xuống dưới, phân tầng máng theo độ phủ để không cắt nhau; nhãn không dùng mã MSG. Các tình huống biên đã phủ: OTP hết hạn ≠ OTP sai, tài khoản bị khóa, email trùng → nhảy sang Login, chưa có xe khi đặt lịch, re-check Pending khi confirm (concurrency), LPR quá 5 giây → nhập tay, walk-in nhập tiền tay, không tài khoản → không gửi email, chặn tự khóa tài khoản mình.

### diagrams/StateTransition.drawio — 1 page, 3 sơ đồ

Booking (Pending → Confirmed → CheckedIn → Completed, nhánh Cancelled — khớp bảng 2.4.1 SRS); Queue (composite "In progress" bao 5 công đoạn, một mũi tên Cancelled với guard [not Completed] — khớp 2.4.2); OTP lifecycle (Active → Verified/Expired — khớp CBR3).

### tools/srs-generator/ — sinh lại file SRS docx

`build.js` chứa toàn bộ nội dung các mục 1, 2, 4–8 + lắp ráp document; `uc1.js`/`uc2.js` chứa dữ liệu 24 đặc tả UC (BR tự đánh số khi render); `helpers.js` là hàm dựng bảng/heading theo docx-js. Sinh lại file:

```
npm install docx
node build.js        # tạo Auto-Wash_SRS_v2.4.docx tại chỗ (tự chèn ERD.png + mockups/*.png nếu có)
```

Muốn sửa nội dung SRS thì sửa các file .js rồi build lại, đừng sửa tay file docx để tránh lệch nguồn.

## Nhật ký công việc đã làm

1. Viết lại SRS v1.0 (tiếng Anh, phạm vi rộng) thành **v2.0 tiếng Việt** theo đúng phạm vi ContextDiagram: bỏ loyalty/tier, voucher, campaign, dashboard, Google OAuth; thêm LPR, Quản lý dịch vụ/tài khoản, Xác nhận/Hủy đặt lịch cho Staff.
2. Review sơ đồ ngữ cảnh, sửa 3 lỗi thành **ContextDiagram_v2**: thêm luồng phản hồi cho Customer, đổi tên 3rd-party LPR API, thêm request plate scan; sau đó bổ sung 3 luồng gộp để phủ đủ UC.
3. Vẽ **UseCaseDiagram** (tổng → tách 5 page → chuẩn hóa include/extend, bỏ functional decomposition, bỏ tiền tố UC, tên tiếng Anh, Login include OTP).
4. Đồng bộ SRS lên **v2.1**: UC1 thêm OTP hai lớp, mục 2.5 thêm ghi chú quan hệ, UC3 thêm điều hướng sang Login, UC12 thêm bước tóm tắt + xác nhận, bảng 2.1 khớp context.
5. Vẽ **ActivityDiagrams + Swimlanes** (17 luồng, qua nhiều vòng: bỏ màu, chuyển ngang, gộp cụm, chống chồng chéo, bỏ mã MSG).
6. Vẽ **StateTransition** (Booking, Queue, OTP).
7. Dọn dẹp: xóa BusinessFlow.drawio (bản luồng cũ bị thay thế).
8. Rà soát SRS lên **v2.2**: kiểm tra toàn bộ 24 UC theo checklist tình huống biên ở bước 5, bổ sung 3 bước kiểm tra + 7 BR/MSG còn thiếu (UC1, UC6, UC11, UC12, UC14, UC15, UC19). Không đổi diagram lần này.
9. **Đồng bộ & hoàn thiện lên v2.3** (theo Auto-Wash_Plan_CaiTien.xlsx): gắn mã UC1–UC24 và chốt 24 UC trên UseCaseDiagram; thêm nhánh transition UC19 vào Activity/Swimlane + ghi chú State; thêm luồng notifications vào Context; bổ sung Glossary (1.4.1) và ghi chú UC không vẽ Activity; vẽ ERD 9 thực thể + chèn vào mục 2.2; tạo mockup wireframe cho cả 23 màn (SVG) + chèn vào mục 4 và Mockups.html; build lại SRS v2.3; dọn v2.1/v2.2 vào legacy.
10. **Bỏ OTP khi đăng nhập → v2.4:** UC1 chỉ dùng định danh + mật khẩu. Đồng bộ toàn bộ: UseCaseDiagram bỏ «include» Login→Verify OTP (còn 3 include); Activity & Swimlane rút gọn luồng "B. Login" (bỏ sinh/gửi/nhập/kiểm tra OTP, nối thẳng kiểm tra mật khẩu → tạo session); StateTransition đổi nguồn OTP thành UC3/UC5/UC8; mockup Đăng nhập bỏ ô OTP; SRS cập nhật UC1, mục 2.5, ET1, MSG4; build lại v2.4; dọn v2.3 vào legacy.

## Việc còn dở / gợi ý bước tiếp theo

- Mockup: đã có 6 màn chính; các màn còn lại (landing page, quản lý dịch vụ, các dialog check-in/check-out/hủy...) nên bổ sung tiếp vào diagrams/mockups/ và map thêm trong build.js.
- ERD.png đang render bằng script; nếu muốn hình chuẩn drawio, mở diagrams/ERD.drawio và File > Export as PNG đè lên ERD.png rồi build lại.
- Các sơ đồ khác (Context/UseCase/Activity/Swimlane/State) trong SRS vẫn ghi [TBU/xem file drawio]: khi cần, export PNG từ drawio và map vào build.js giống ERD/mockup.
- Khi đổi nghiệp vụ, sửa theo thứ tự: ContextDiagram → UseCaseDiagram → Activity/Swimlane/StateTransition → SRS (qua srs-generator) để giữ đồng bộ.
