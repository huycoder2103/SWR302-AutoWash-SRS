const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat,
  TableOfContents, PageBreak, PageNumber, Header, Footer, TabStopType, TabStopPosition,
  ImageRun,
} = require('docx');
const path = require('path');
const H = require('./helpers');
const uc1 = require('./uc1');
const uc2 = require('./uc2');
const { t, p, h1, h2, h3, table, spacer, bullet, renderUC, CW } = H;

const brCounter = { n: 1 };
const children = [];

function figart(sub, file, w, h, caption) {
  const fp = path.join(__dirname, '..', '..', 'diagrams', 'diagram picture', sub, file);
  const out = [];
  if (fs.existsSync(fp)) {
    out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
      children: [ new ImageRun({ type: 'png', data: fs.readFileSync(fp), transformation: { width: w, height: h } }) ] }));
    out.push(p(caption, { run: { italics: true, size: 18 } }));
  } else { out.push(p(caption + ' [thiếu ảnh: ' + file + ']', { run: { italics: true } })); }
  return out;
}

/* ---------------- Title page ---------------- */
children.push(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2000, after: 240 }, children: [t('FPT University — SWP391', { bold: true, size: 28 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [t('TÀI LIỆU ĐẶC TẢ YÊU CẦU PHẦN MỀM', { bold: true, size: 40 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [t('(Software Requirements Specification)', { italics: true, size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [t('Dành cho', { size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 }, children: [t('Auto-Wash — Hệ thống quản lý trung tâm rửa xe tự động', { bold: true, size: 32 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [t('Phiên bản: 2.8 (chèn đầy đủ mọi sơ đồ đã vẽ: thêm Activity, Swimlane, Data Flow vào mục 2.3; sửa văn bản đăng nhập không OTP)', { size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [t('TP. Hồ Chí Minh, tháng 7 năm 2026', { size: 24 })] }),
  new Paragraph({ children: [new PageBreak()] }),
);

/* ---------------- Approval + revision ---------------- */
children.push(h1('Trang phê duyệt'));
children.push(p('Chữ ký trên tài liệu này thể hiện sự đồng thuận đối với tài liệu "Đặc tả yêu cầu phần mềm Auto-Wash".'));
children.push(table(
  ['Vai trò', 'Họ và tên', 'Chữ ký', 'Ngày'],
  [
    ['Người lập', 'HoangHuy', '____________________', '____ / ____ / 2026'],
    ['Người rà soát', '', '____________________', '____ / ____ / 2026'],
    ['Người phê duyệt', '', '____________________', '____ / ____ / 2026'],
  ],
  [2000, 2400, 2626, 2000],
));
children.push(spacer());
children.push(h1('Lịch sử sửa đổi'));
children.push(table(
  ['Ngày', 'Phiên bản', 'Tác giả', 'Mô tả thay đổi'],
  [
    ['06/06/2026', '1.0', 'HoangHuy', 'Khởi tạo — SRS đầy đủ cho Auto-Wash v2'],
    ['08/07/2026', '2.0', 'HoangHuy', 'Viết lại toàn bộ theo Sơ đồ ngữ cảnh mới: giới hạn phạm vi trong 5 dịch vụ (Account, Booking, Queue, Notification, LPR); loại bỏ chương trình khách hàng thân thiết (điểm, hạng thành viên), voucher, chiến dịch khuyến mãi, dashboard quản trị và Google OAuth; bổ sung tích hợp Camera LPR và các chức năng Quản lý dịch vụ, Quản lý tài khoản, Xác nhận/Hủy đặt lịch cho nhân viên. Đồng bộ theo ContextDiagram_v2: đổi tên hệ thống ngoài thành "3rd-party LPR API", bổ sung luồng request plate scan và các luồng phản hồi cho khách hàng.'],
    ['11/07/2026', '2.1', 'HoangHuy', 'Đồng bộ theo Use Case Diagram chuẩn hóa: UC1 bổ sung xác thực OTP hai lớp khi đăng nhập (luồng, BR, MSG, ET cập nhật theo); mục 2.5 bổ sung ghi chú quan hệ «include»/«extend» cho từng phân hệ. Sơ đồ ngữ cảnh bổ sung luồng update queue status, view vehicles & services, manage account (bảng 2.1 cập nhật theo). Đồng bộ theo Activity/Swimlane Diagrams: UC3 bổ sung điều hướng sang Đăng nhập khi email đã tồn tại; UC12 bổ sung bước xem tóm tắt và xác nhận trước khi lưu.'],
    ['12/07/2026', '2.2', 'HoangHuy', 'Rà soát và bổ sung nhánh lỗi/ngoại lệ còn thiếu cho phần 3 (24 UC), đối chiếu với các tình huống biên đã liệt kê ở Activity/Swimlane Diagrams: UC1 chặn đăng nhập khi tài khoản bị vô hiệu hóa (MSG26); UC6 hiển thị thông báo khi cố tự khóa tài khoản (MSG27); UC12 kiểm tra và chặn tạo đặt lịch khi khách hàng chưa có xe (MSG28); UC14 xử lý tranh chấp trạng thái đồng thời khi xác nhận booking (MSG29); UC15 hiển thị thông báo khi hủy booking sai trạng thái (MSG30); UC19 hiển thị thông báo khi chuyển trạng thái không hợp lệ (MSG31); UC11 hiển thị thông báo khi giá/thời lượng dịch vụ không hợp lệ (MSG32). Phụ lục 8.1 bổ sung MSG26–MSG32.'],
    ['15/07/2026', '2.3', 'HoangHuy', 'Đồng bộ toàn bộ sơ đồ với SRS: (1) Use Case Diagram gắn mã UC1–UC24 và chú thích «Track queue progress» là nhánh con của UC13 (chốt đúng 24 UC); (2) Activity & Swimlane bổ sung nhánh kiểm tra chuyển trạng thái hàng chờ hợp lệ cho UC19 (khớp MSG31); (3) Sơ đồ ngữ cảnh bổ sung luồng view & mark notifications (UC23/UC24) cho Customer; (4) State Transition thêm ghi chú ràng buộc thứ tự công đoạn (UC19). Bổ sung mục 1.4.1 Bảng thuật ngữ nghiệp vụ chuẩn hóa (Glossary) và ghi chú các UC chủ đích không vẽ Activity (UC2, UC10, UC16).'],
    ['16/07/2026', '2.4', 'HoangHuy', 'Thay đổi nghiệp vụ: bỏ xác thực OTP khi đăng nhập — UC1 chỉ dùng định danh + mật khẩu. Cập nhật đồng bộ: Use Case Diagram bỏ «include» Login→Verify OTP; Activity & Swimlane rút gọn luồng B. Login (bỏ các bước sinh/gửi/nhập/kiểm tra OTP); State Transition đổi nguồn OTP thành UC3/UC5/UC8; mockup Đăng nhập bỏ ô OTP; Phụ lục ET1 và MSG4 bỏ tham chiếu UC1. OTP vẫn dùng cho UC3 (đăng ký), UC5 (đổi mật khẩu), UC8 (đăng ký xe).'],
    ['16/07/2026', '2.5', 'HoangHuy', 'Hoàn thiện Đặc tả Use Case (mục 3) cho cả 24 UC: đổi "Luồng hoạt động" thành "Luồng sự kiện chính" và bổ sung mục "Luồng thay thế / ngoại lệ" (tổng hợp các điều kiện lỗi/từ chối kèm cách xử lý) bên cạnh bảng Quy tắc nghiệp vụ. Không đổi nghiệp vụ; chỉ làm đầy đủ cấu trúc đặc tả UC.'],
    ['16/07/2026', '2.6', 'HoangHuy', 'Bổ sung mục 4.1 Sơ đồ luồng màn hình (Dialog Map): hai sơ đồ Khách hàng (Mobile Web) và Nhân viên (Desktop) thể hiện điều hướng giữa 23 màn; danh mục mockup dời xuống 4.2–4.6. Nguồn: diagrams/DialogMap_Customer.drawio, diagrams/DialogMap_Staff.drawio.'],
    ['17/07/2026', '2.7', 'HoangHuy', 'Chèn ảnh sơ đồ (xuất từ drawio) vào mục 2: Hình 1 Sơ đồ ngữ cảnh, Hình 2 ERD, Hình 3 Chuyển trạng thái (Booking/Queue/OTP), Hình 5–9 Use Case 5 phân hệ; xóa toàn bộ [TBU] hình trong mục 2.'],
    ['17/07/2026', '2.8', 'HoangHuy', 'Chèn nốt các sơ đồ còn lại (xuất từ drawio): mục 2.3 thêm 2.3.4 Activity (5 hình), 2.3.5 Swimlane (5 hình), 2.3.6 Data Flow (1 hình). Sửa văn bản 2.3.1: đăng nhập không dùng OTP. SRS nay thể hiện đầy đủ toàn bộ sơ đồ.'],
  ],
  [1400, 1200, 1400, 5026],
));
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- TOC ---------------- */
children.push(h1('Mục lục'));
children.push(new TableOfContents('Mục lục', { hyperlink: true, headingStyleRange: '1-3' }));
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- 1. Giới thiệu ---------------- */
children.push(h1('1. Giới thiệu'));
children.push(h2('1.1 Mục đích'));
children.push(p('Tài liệu Đặc tả yêu cầu phần mềm (SRS) này định nghĩa các yêu cầu chức năng và phi chức năng cho Hệ thống quản lý trung tâm rửa xe tự động Auto-Wash, với phạm vi được xác định theo Sơ đồ ngữ cảnh (Context Diagram) phiên bản mới nhất. Tài liệu nhằm:'));
children.push(bullet('Xác định phạm vi mục tiêu nghiệp vụ, chức năng nghiệp vụ và các đơn vị liên quan.'));
children.push(bullet('Nhận diện các quy trình nghiệp vụ mà giải pháp phải hỗ trợ.'));
children.push(bullet('Tạo cách hiểu thống nhất về yêu cầu chức năng cho tất cả các bên tham gia.'));
children.push(bullet('Làm cơ sở xây dựng kiểm thử chấp nhận (acceptance test) nhằm xác nhận giải pháp đáp ứng yêu cầu.'));

children.push(h2('1.2 Tổng quan'));
children.push(p('Auto-Wash là hệ thống quản lý rửa xe trên nền web dành cho một trung tâm rửa xe tự động. Theo Sơ đồ ngữ cảnh, hệ thống gồm 5 dịch vụ nội bộ: Account Service (tài khoản), Booking Service (đặt lịch), Queue Service (hàng chờ), Notification Service (thông báo) và LPR Service (nhận dạng biển số). Hệ thống hỗ trợ:'));
children.push(bullet('Khách hàng tự phục vụ qua Mobile Web: đăng ký tài khoản, đăng nhập, đăng ký xe, đặt lịch rửa xe và theo dõi lịch đặt.'));
children.push(bullet('Nhân viên vận hành: check-in/check-out khách hàng, quét biển số bằng camera LPR, theo dõi hàng chờ, xác nhận/hủy đặt lịch, quản lý tài khoản người dùng và quản lý danh mục dịch vụ.'));
children.push(bullet('Tự động hóa: gửi email OTP, email xác nhận đặt lịch và email thông báo rửa xe hoàn tất qua máy chủ Email SMTP; nhận dạng biển số qua Camera LPR kết hợp 3rd-party LPR API.'));
children.push(spacer());
children.push(p('Tác nhân và vai trò:', { run: { bold: true } }));
children.push(table(
  ['Tác nhân', 'Vai trò & trách nhiệm'],
  [
    ['Nhân viên (Staff)', 'Vận hành hàng ngày: check-in khách hàng, quét biển số, theo dõi và cập nhật hàng chờ, check-out, xác nhận và hủy đặt lịch, quản lý tài khoản người dùng, quản lý danh mục dịch vụ.'],
    ['Khách hàng (Customer)', 'Đăng ký tài khoản, đăng nhập, quản lý xe, tạo đặt lịch và theo dõi lịch đặt qua Mobile Web; nhận thông báo qua email.'],
    ['Hệ thống (System)', 'Tác vụ tự động: gửi email OTP, email xác nhận đặt lịch, email thông báo hoàn tất; tiếp nhận kết quả nhận dạng biển số từ 3rd-party LPR API.'],
  ],
  [2400, 6626],
));
children.push(spacer());
children.push(p('Hệ thống ngoài (external systems):', { run: { bold: true } }));
children.push(table(
  ['Hệ thống ngoài', 'Vai trò'],
  [
    ['Email SMTP', 'Máy chủ gửi thư: chuyển email OTP, email xác nhận đặt lịch và email thông báo rửa xe hoàn tất đến khách hàng.'],
    ['Camera LPR', 'Thiết bị camera đặt tại trung tâm, chụp ảnh biển số xe theo yêu cầu của hệ thống (request plate scan).'],
    ['3rd-party LPR API', 'Dịch vụ nhận dạng ký tự bên ngoài: nhận ảnh biển số từ Camera LPR, trích xuất chuỗi biển số và trả kết quả về LPR Service của hệ thống.'],
  ],
  [2400, 6626],
));

children.push(h2('1.3 Đối tượng đọc và gợi ý đọc'));
children.push(bullet('Nhóm phát triển: thiết kế, lập trình, kiểm thử đơn vị và kiểm thử tích hợp dựa trên đặc tả này.'));
children.push(bullet('Nhóm QA/UAT: xây dựng và thực thi các ca kiểm thử chấp nhận.'));
children.push(bullet('Bên liên quan nghiệp vụ / Giảng viên hướng dẫn: rà soát và phê duyệt yêu cầu.'));
children.push(bullet('Nhóm thiết kế UI/UX: thiết kế màn hình bám sát các use case và mục mockup.'));

children.push(h2('1.4 Thuật ngữ viết tắt'));
children.push(table(
  ['Từ viết tắt', 'Giải nghĩa'],
  [
    ['SRS', 'Software Requirements Specification — Đặc tả yêu cầu phần mềm'],
    ['UC', 'Use Case — Ca sử dụng'],
    ['BR', 'Business Rule — Quy tắc nghiệp vụ'],
    ['CBR', 'Common Business Rule — Quy tắc nghiệp vụ chung'],
    ['MSG', 'System Message — Thông điệp hệ thống hiển thị cho người dùng'],
    ['ET', 'Email Template — Mẫu email'],
    ['OTP', 'One-Time Password — Mã xác thực dùng một lần (6 chữ số)'],
    ['LPR', 'License Plate Recognition — Nhận dạng biển số xe'],
    ['VND', 'Đồng Việt Nam (đơn vị tiền tệ)'],
    ['SMTP', 'Simple Mail Transfer Protocol — Giao thức gửi thư điện tử'],
    ['API', 'Application Programming Interface — Giao diện lập trình ứng dụng'],
    ['DB', 'Database — Cơ sở dữ liệu (PostgreSQL qua Supabase)'],
    ['FIFO', 'First In First Out — Đến trước phục vụ trước'],
    ['TBU', 'To Be Updated — Sẽ cập nhật sau'],
  ],
  [2000, 7026],
));

children.push(h3('1.4.1 Bảng thuật ngữ nghiệp vụ chuẩn hóa (Glossary)'));
children.push(p('Mỗi khái niệm dùng đúng một thuật ngữ thống nhất trên toàn bộ sơ đồ và tài liệu, tránh dùng nhiều tên cho cùng một đối tượng.'));
children.push(table(
  ['Thuật ngữ chuẩn', 'Định nghĩa', 'Các tên khác đã hợp nhất (không dùng)'],
  [
    ['Đăng ký tài khoản (Register account)', 'Tạo tài khoản khách hàng mới, xác thực OTP qua email — UC3.', 'Register'],
    ['Đặt lịch (Booking)', 'Yêu cầu rửa xe do khách tạo, có trạng thái Pending/Confirmed/CheckedIn/Completed/Cancelled.', 'Đơn đặt, lịch hẹn'],
    ['Mục hàng chờ (Queue entry)', 'Bản ghi theo dõi một xe qua các công đoạn rửa trong ngày; liên kết tùy chọn với một Booking.', 'Queue, hàng đợi'],
    ['Cập nhật hàng chờ (Update queue)', 'Nhân viên chỉnh trạng thái/ghi chú một mục hàng chờ — UC20.', 'Update queue entry, update queue status'],
    ['Chuyển trạng thái rửa (Advance wash status)', 'Chuyển mục hàng chờ sang công đoạn kế tiếp theo đúng thứ tự — UC19.', 'Advance stage'],
    ['Theo dõi tiến trình hàng chờ (Track queue progress)', 'Nhánh con của UC13 (Xem đặt lịch) cho phép khách xem vị trí và công đoạn hiện tại; không phải UC độc lập.', '(được đếm là một phần UC13)'],
    ['Công đoạn rửa (Wash stage)', 'Waiting → LPR_Scan → Washing → Addon_Processing → Drying → Completed.', 'Bước rửa, giai đoạn'],
  ],
  [2600, 4026, 2400],
));
children.push(spacer());

children.push(h2('1.5 Tài liệu tham khảo'));
children.push(bullet('Sơ đồ ngữ cảnh Auto-Wash (ContextDiagram.drawio) — căn cứ xác định phạm vi của tài liệu này.'));
children.push(bullet('FA Management Requirement Specification (FPT Fresher Academy, 2018) — dùng làm mẫu định dạng SRS.'));
children.push(bullet('Auto-Wash Backend — ASP.NET Core 8, Entity Framework Core, PostgreSQL (Supabase).'));
children.push(bullet('Auto-Wash Frontend Mobile Web — React + Vite.'));
children.push(bullet('IEEE 830-1998: Recommended Practice for Software Requirements Specifications.'));
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- 2. Yêu cầu tổng quát ---------------- */
children.push(h1('2. Yêu cầu tổng quát'));
children.push(p('Phần này mô tả tổng quan các chức năng của hệ thống bằng sơ đồ: loại người dùng, quyền được cấp và trình tự hoàn thành các luồng nghiệp vụ. Đặc tả chi tiết xem tại Phần 3.'));

children.push(h2('2.1 Sơ đồ ngữ cảnh (Context Diagram)'));
children.push(...figart('Context', 'ContextDiagram_v2.drawio.png', 600, 306, 'Hình 1: Sơ đồ ngữ cảnh hệ thống Auto-Wash — nguồn: diagrams/ContextDiagram_v2.drawio.'));
children.push(p('Sơ đồ ngữ cảnh xác định ranh giới hệ thống, các tác nhân và hệ thống ngoài cùng các luồng tương tác chính:'));
children.push(table(
  ['Phần tử', 'Loại', 'Mô tả'],
  [
    ['Auto-Wash System', 'Hệ thống', 'Gồm 5 dịch vụ nội bộ: Account Service, Booking Service, Queue Service, Notification Service, LPR Service.'],
    ['Mobile Web', 'Kênh truy cập', 'Ứng dụng web trên di động — kênh tương tác của Khách hàng, giao tiếp với hệ thống qua API request/response.'],
    ['Nhân viên (Staff)', 'Tác nhân', 'Tương tác trực tiếp với hệ thống: check-in, quét biển số, check-out, xem hàng chờ, quản lý tài khoản, quản lý dịch vụ, xác nhận/hủy đặt lịch.'],
    ['Khách hàng (Customer)', 'Tác nhân', 'Tương tác qua Mobile Web: đăng ký, đăng nhập, đăng ký xe, tạo đặt lịch, xem lịch đặt; nhận thông báo qua email.'],
    ['Camera LPR', 'Hệ thống ngoài', 'Chụp ảnh biển số và gửi đến 3rd-party LPR API.'],
    ['3rd-party LPR API', 'Hệ thống ngoài', 'Trích xuất chuỗi biển số từ ảnh và trả kết quả về hệ thống.'],
    ['Email SMTP', 'Hệ thống ngoài', 'Nhận yêu cầu gửi email từ hệ thống (OTP, xác nhận đặt lịch, thông báo hoàn tất) và chuyển đến khách hàng.'],
  ],
  [2300, 1600, 5126],
));
children.push(spacer());
children.push(p('Các luồng tương tác trên sơ đồ:', { run: { bold: true } }));
children.push(table(
  ['Nguồn', 'Đích', 'Luồng dữ liệu'],
  [
    ['Nhân viên', 'Hệ thống', 'check in customer; scan license plate; check out customer; view queue; update queue status (chuyển trạng thái/cập nhật/hủy lượt); manage user accounts; manage services; confirm booking; cancel booking'],
    ['Hệ thống', 'Nhân viên', 'return queue list; return booking details; return user list'],
    ['Khách hàng', 'Mobile Web', 'Register; Login; register vehicle; view vehicles & services; create booking; view bookings; manage account (profile, password)'],
    ['Mobile Web', 'Hệ thống', 'send API request'],
    ['Hệ thống', 'Mobile Web', 'return API response'],
    ['Mobile Web', 'Khách hàng', 'return account info; return vehicle list; return booking list & status'],
    ['Hệ thống', 'Camera LPR', 'request plate scan (yêu cầu chụp ảnh biển số)'],
    ['Camera LPR', '3rd-party LPR API', 'capture plate image (ảnh biển số)'],
    ['3rd-party LPR API', 'Hệ thống', 'extract plate text (chuỗi biển số)'],
    ['Hệ thống', 'Email SMTP', 'send OTP email; send booking confirmation; send wash-complete notification'],
    ['Email SMTP', 'Khách hàng', 'deliver notification (chuyển email đến khách hàng)'],
  ],
  [2100, 2100, 4826],
));

children.push(h2('2.2 Sơ đồ thực thể (Entity Relationship Diagram)'));
const erdPath = path.join(__dirname, '..', '..', 'diagrams', 'diagram picture', 'ERD', 'ERD.png');
if (fs.existsSync(erdPath)) {
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [ new ImageRun({ type: 'png', data: fs.readFileSync(erdPath), transformation: { width: 540, height: 400 } }) ],
  }));
  children.push(p('Hình 2: Sơ đồ thực thể của hệ thống Auto-Wash (9 thực thể) — nguồn: diagrams/ERD.drawio.', { run: { italics: true } }));
} else {
  children.push(p('Hình 2: Sơ đồ thực thể của hệ thống Auto-Wash [xem diagrams/ERD.drawio]', { run: { italics: true } }));
}
children.push(p('Các thực thể chính và quan hệ:'));
children.push(table(
  ['Thực thể', 'Trường chính', 'Quan hệ'],
  [
    ['Account', 'AccountId, Email, Phone, PasswordHash, FullName, Role (1=Staff, 2=Customer), IsActive', '1:1 với Customer (khi Role = Customer)'],
    ['Customer', 'CustomerId, AccountId, TotalVisits, TotalSpend, LastVisitAt', '1:N Vehicle, Booking, Notification, Queue'],
    ['Vehicle', 'VehicleId, CustomerId, LicensePlate, Brand, Name, RegisteredAt', 'N:1 Customer; 1:N Booking, Queue'],
    ['Booking', 'BookingId, CustomerId, VehicleId, ScheduledAt, Status (1–5), BasePrice, FinalPrice', 'N:1 Customer, Vehicle; 1:N BookingService, Queue'],
    ['BookingService', 'BookingServiceId, BookingId, ServiceId, PriceSnapshot', 'N:1 Booking, Service (bảng trung gian)'],
    ['Service', 'ServiceId, ServiceName, Category (1=Cơ bản, 2=Cao cấp, 3=Đặc biệt, 4=AddOn), BasePrice, EstimatedMinutes, IsAddOn, IsActive', '1:N BookingService'],
    ['Queue', 'QueueId, BookingId, VehicleId, CustomerId, LicensePlate, Status (text), Position, CheckInAt, StartedAt, CompletedAt, StaffNote', 'N:1 Booking, Vehicle, Customer'],
    ['Notification', 'NotificationId, CustomerId, Title, Message, Type, IsRead, CreatedAt', 'N:1 Customer'],
    ['OtpVerification', 'OtpId, Email, Code, ExpiresAt, IsUsed, CreatedAt', 'Tạm thời — không có quan hệ khóa ngoại'],
  ],
  [1700, 4326, 3000],
));

children.push(h2('2.3 Luồng nghiệp vụ (Workflow)'));
children.push(h3('2.3.1 Quản lý tài khoản'));
children.push(p('Khách hàng mới: truy cập Mobile Web → nhấn Đăng ký → nhập email → nhận email OTP (qua Email SMTP) → điền Họ tên, Số điện thoại, Mật khẩu, mã OTP → tài khoản được tạo → chuyển đến trang chủ.'));
children.push(p('Khách hàng hiện hữu: nhập email/số điện thoại + mật khẩu → hệ thống xác thực → session được tạo → chuyển đến trang chủ khách hàng (đăng nhập không dùng OTP).'));
children.push(p('Nhân viên: nhập email + mật khẩu → hệ thống xác thực → session được tạo → chuyển đến màn hình Quản lý hàng chờ. Nhân viên có thể mở màn hình Quản lý tài khoản để xem danh sách người dùng và kích hoạt/vô hiệu hóa tài khoản.'));
children.push(h3('2.3.2 Quản lý đặt lịch'));
children.push(p('Khách hàng đăng nhập → đăng ký xe (xác thực OTP) → xem trang dịch vụ → chọn dịch vụ chính + dịch vụ bổ sung (tùy chọn) → chọn ngày giờ → xem lại giá → xác nhận đặt lịch → hệ thống gửi email xác nhận đặt lịch → nhân viên xác nhận booking → khách hàng đến trung tâm đúng lịch hẹn. Nhân viên có thể hủy booking khi khách hàng yêu cầu.'));
children.push(h3('2.3.3 Quản lý hàng chờ'));
children.push(p('Xe đến trung tâm → nhân viên nhấn quét biển số, hệ thống gửi yêu cầu chụp ảnh (request plate scan) đến Camera LPR → Camera chụp ảnh và gửi đến 3rd-party LPR API → 3rd-party LPR API trích xuất chuỗi biển số và trả về hệ thống (hoặc nhân viên nhập biển số thủ công) → hệ thống tự động phát hiện booking nếu có → tạo mục Queue (Status: Đang chờ) → nhân viên chuyển trạng thái qua các công đoạn: Quét LPR → Đang rửa → Xử lý add-on → Sấy khô → nhân viên check-out → thống kê khách hàng được cập nhật → hệ thống gửi email thông báo rửa xe hoàn tất.'));

children.push(h3('2.3.4 Sơ đồ hoạt động (Activity Diagram)'));
children.push(p('Sơ đồ hoạt động mô tả chi tiết luồng xử lý theo từng phân hệ (vẽ trục ngang; nguồn: diagrams/ActivityDiagrams.drawio).'));
children.push(...figart('Activity', 'ActivityDiagrams-1. Account.drawio.png', 600, 168, 'Hình 10: Activity — Quản lý tài khoản.'));
children.push(...figart('Activity', 'ActivityDiagrams-2. Vehicle.drawio.png', 600, 179, 'Hình 11: Activity — Quản lý phương tiện.'));
children.push(...figart('Activity', 'ActivityDiagrams-3. Booking.drawio.png', 600, 250, 'Hình 12: Activity — Quản lý dịch vụ & đặt lịch.'));
children.push(...figart('Activity', 'ActivityDiagrams-4. Queue.drawio.png', 600, 169, 'Hình 13: Activity — Quản lý hàng chờ.'));
children.push(...figart('Activity', 'ActivityDiagrams-5. Administration - Notifications.drawio.png', 600, 247, 'Hình 14: Activity — Quản trị & Thông báo.'));
children.push(h3('2.3.5 Sơ đồ phân làn (Swimlane Diagram)'));
children.push(p('Cùng các luồng nhưng phân làn theo tác nhân: Customer / Staff / Auto-Wash System / Email SMTP (nguồn: diagrams/Swimlanes.drawio).'));
children.push(...figart('SwimLane', 'Swimlanes-1. Account.drawio.png', 600, 222, 'Hình 15: Swimlane — Quản lý tài khoản.'));
children.push(...figart('SwimLane', 'Swimlanes-2. Vehicle.drawio.png', 600, 184, 'Hình 16: Swimlane — Quản lý phương tiện.'));
children.push(...figart('SwimLane', 'Swimlanes-3. Booking.drawio.png', 600, 199, 'Hình 17: Swimlane — Quản lý dịch vụ & đặt lịch.'));
children.push(...figart('SwimLane', 'Swimlanes-4. Queue.drawio.png', 600, 159, 'Hình 18: Swimlane — Quản lý hàng chờ.'));
children.push(...figart('SwimLane', 'Swimlanes-5. Administration - Notifications.drawio.png', 600, 156, 'Hình 19: Swimlane — Quản trị & Thông báo.'));
children.push(h3('2.3.6 Sơ đồ luồng dữ liệu (Data Flow Diagram — Mức 1)'));
children.push(...figart('DataFlow', 'DataFlowDiagram_Level1.drawio.png', 600, 301, 'Hình 20: Sơ đồ luồng dữ liệu mức 1 (DFD Level 1).'));
children.push(h2('2.4 Sơ đồ chuyển trạng thái'));
children.push(...figart('StateTransition', 'StateTransition.drawio.png', 600, 355, 'Hình 3: Sơ đồ chuyển trạng thái — Booking, Queue entry (composite In progress), OTP lifecycle (nguồn: diagrams/StateTransition.drawio).'));
children.push(h3('2.4.1 Trạng thái đặt lịch (Booking)'));
children.push(p('Bảng chi tiết chuyển trạng thái Booking (phần Booking của Hình 3):', { run: { italics: true } }));
children.push(table(
  ['Trạng thái hiện tại', 'Sự kiện chuyển', 'Trạng thái kế tiếp'],
  [
    ['(chưa có)', 'Khách hàng tạo đặt lịch (UC12)', 'Chờ xác nhận (1)'],
    ['Chờ xác nhận', 'Nhân viên xác nhận đặt lịch (UC14)', 'Đã xác nhận (2)'],
    ['Đã xác nhận', 'Xe đến; nhân viên check-in (UC17)', 'Đã check-in (3)'],
    ['Đã check-in', 'Nhân viên check-out (UC21)', 'Hoàn tất (4)'],
    ['Chờ xác nhận / Đã xác nhận', 'Nhân viên hủy đặt lịch (UC15) hoặc hủy lượt hàng chờ (UC22)', 'Đã hủy (5)'],
  ],
  [2600, 4026, 2400],
));
children.push(h3('2.4.2 Trạng thái hàng chờ (Queue)'));
children.push(p('Bảng chi tiết chuyển trạng thái hàng chờ (phần Queue của Hình 3):', { run: { italics: true } }));
children.push(table(
  ['Trạng thái hiện tại', 'Sự kiện / điều kiện chuyển', 'Trạng thái kế tiếp'],
  [
    ['(chưa có)', 'Check-in khách hàng (UC17)', 'Đang chờ (Waiting)'],
    ['Đang chờ', 'Nhân viên chuyển tiếp (UC19)', 'Quét LPR (LPR_Scan)'],
    ['Quét LPR', 'Nhân viên chuyển tiếp (UC19)', 'Đang rửa (Washing)'],
    ['Đang rửa', 'Nhân viên chuyển tiếp (UC19)', 'Xử lý add-on (Addon_Processing)'],
    ['Xử lý add-on', 'Nhân viên chuyển tiếp (UC19)', 'Sấy khô (Drying)'],
    ['Sấy khô', 'Nhân viên check-out (UC21)', 'Hoàn tất (Completed)'],
    ['Bất kỳ (trừ Hoàn tất)', 'Nhân viên hủy lượt (UC22)', 'Đã hủy (Cancelled)'],
  ],
  [2600, 3726, 2700],
));

children.push(h2('2.5 Sơ đồ Use Case'));
children.push(p('Sơ đồ Use Case gồm 24 UC (UC1–UC24) tách theo 5 phân hệ, mỗi bong bóng gắn mã UC khớp Phần 3. Sub-use-case "Xác thực OTP qua email" là bước dùng chung («include»), không được đánh số. "Track queue progress" là nhánh con của UC13 (thể hiện bằng «extend»), không tính là UC độc lập.', { run: { italics: true } }));
children.push(p('Ghi chú độ phủ: UC2 (Đăng xuất), UC10 (Xem dịch vụ) và UC16 (Xem hàng chờ hôm nay) là thao tác đọc/tầm thường nên chủ đích không vẽ Activity/Swimlane riêng; các UC này vẫn có mặt đầy đủ trên Use Case Diagram và Sơ đồ ngữ cảnh.', { run: { italics: true } }));
children.push(h3('2.5.1 Chung & Quản lý tài khoản'));
children.push(...figart('UseCase', 'UseCaseDiagram-1. General & Account Management.drawio.png', 470, 304, 'Hình 5: Sơ đồ Use Case — Chung & Quản lý tài khoản.'));
children.push(table(
  ['#', 'Tên UC', 'Mô tả'],
  [
    ['UC1', 'Đăng nhập', 'Xác thực bằng email/số điện thoại + mật khẩu.'],
    ['UC2', 'Đăng xuất', 'Xóa session và chuyển về màn hình đăng nhập.'],
    ['UC3', 'Đăng ký tài khoản', 'Tạo tài khoản khách hàng mới với xác thực OTP qua email.'],
    ['UC4', 'Cập nhật hồ sơ', 'Khách hàng cập nhật họ tên và số điện thoại.'],
    ['UC5', 'Đổi mật khẩu', 'Khách hàng đổi mật khẩu với xác thực OTP qua email.'],
    ['UC6', 'Quản lý tài khoản người dùng', 'Nhân viên xem danh sách người dùng, kích hoạt/vô hiệu hóa tài khoản.'],
  ],
  [800, 2800, 5426],
));
children.push(p('Quan hệ: UC3, UC5 «include» sub-UC "Xác thực OTP qua email" (dùng chung, gửi mã qua Email SMTP). UC1 (Đăng nhập) chỉ dùng mật khẩu, không «include» OTP.', { run: { italics: true } }));
children.push(h3('2.5.2 Quản lý phương tiện'));
children.push(...figart('UseCase', 'UseCaseDiagram-2. Vehicle Management.drawio.png', 470, 209, 'Hình 6: Sơ đồ Use Case — Quản lý phương tiện.'));
children.push(table(
  ['#', 'Tên UC', 'Mô tả'],
  [
    ['UC7', 'Xem danh sách xe', 'Liệt kê tất cả biển số đã đăng ký của khách hàng.'],
    ['UC8', 'Đăng ký xe', 'Thêm xe mới với biển số được xác thực OTP.'],
    ['UC9', 'Xóa xe', 'Xóa xe chưa có lịch sử đặt lịch.'],
  ],
  [800, 2800, 5426],
));
children.push(p('Quan hệ: UC8 «include» "Xác thực OTP qua email". UC9 «extend» UC7 — xóa xe là thao tác tùy chọn phát sinh từ danh sách xe.', { run: { italics: true } }));
children.push(h3('2.5.3 Quản lý dịch vụ & đặt lịch'));
children.push(...figart('UseCase', 'UseCaseDiagram-3. Service & Booking Management.drawio.png', 470, 280, 'Hình 7: Sơ đồ Use Case — Quản lý dịch vụ & đặt lịch.'));
children.push(table(
  ['#', 'Tên UC', 'Mô tả'],
  [
    ['UC10', 'Xem dịch vụ', 'Khách hàng xem các dịch vụ rửa xe và add-on kèm giá.'],
    ['UC11', 'Quản lý dịch vụ', 'Nhân viên thêm/sửa/kích hoạt/vô hiệu hóa dịch vụ.'],
    ['UC12', 'Tạo đặt lịch', 'Khách hàng đặt lịch rửa xe: chọn xe, dịch vụ, ngày giờ.'],
    ['UC13', 'Xem đặt lịch', 'Khách hàng xem booking đang hoạt động và lịch sử, kèm tiến độ hàng chờ.'],
    ['UC14', 'Xác nhận đặt lịch', 'Nhân viên xác nhận booking đang chờ.'],
    ['UC15', 'Hủy đặt lịch', 'Nhân viên hủy booking theo yêu cầu.'],
  ],
  [800, 2800, 5426],
));
children.push(p('Quan hệ: sub-UC "Theo dõi tiến độ hàng chờ" «extend» UC13 — chỉ xuất hiện khi có booking đang hoạt động. UC12 gửi email xác nhận qua Email SMTP.', { run: { italics: true } }));
children.push(h3('2.5.4 Quản lý hàng chờ'));
children.push(...figart('UseCase', 'UseCaseDiagram-4. Queue Management.drawio.png', 470, 350, 'Hình 8: Sơ đồ Use Case — Quản lý hàng chờ.'));
children.push(table(
  ['#', 'Tên UC', 'Mô tả'],
  [
    ['UC16', 'Xem hàng chờ hôm nay', 'Nhân viên xem danh sách hàng chờ hợp nhất của ngày hiện tại.'],
    ['UC17', 'Check-in khách hàng', 'Đưa xe vào hàng chờ khi xe đến; tự động phát hiện booking.'],
    ['UC18', 'Quét biển số (LPR)', 'Nhận dạng biển số tự động qua Camera LPR và 3rd-party LPR API.'],
    ['UC19', 'Chuyển trạng thái rửa xe', 'Chuyển xe sang công đoạn rửa kế tiếp.'],
    ['UC20', 'Cập nhật hàng chờ', 'Cập nhật thủ công trạng thái hoặc thêm ghi chú.'],
    ['UC21', 'Check-out khách hàng', 'Hoàn tất dịch vụ, cập nhật thống kê, gửi thông báo.'],
    ['UC22', 'Hủy lượt hàng chờ', 'Hủy một xe khỏi hàng chờ.'],
  ],
  [800, 2800, 5426],
));
children.push(p('Quan hệ: UC18 «extend» UC17 — quét LPR là phương án tùy chọn thay cho nhập biển số thủ công, kết nối Camera LPR và 3rd-party LPR API. UC21 gửi email hoàn tất qua Email SMTP.', { run: { italics: true } }));
children.push(h3('2.5.5 Thông báo'));
children.push(...figart('UseCase', 'UseCaseDiagram-5. Notification.drawio.png', 470, 249, 'Hình 9: Sơ đồ Use Case — Thông báo.'));
children.push(table(
  ['#', 'Tên UC', 'Mô tả'],
  [
    ['UC23', 'Xem thông báo', 'Khách hàng đọc tất cả thông báo hệ thống.'],
    ['UC24', 'Đánh dấu thông báo đã đọc', 'Đánh dấu một thông báo cụ thể là đã đọc.'],
  ],
  [800, 2800, 5426],
));
children.push(p('Quan hệ: UC24 «extend» UC23 — đánh dấu đã đọc phát sinh khi khách hàng mở một thông báo từ danh sách.', { run: { italics: true } }));

children.push(h2('2.6 Ma trận phân quyền'));
children.push(p('Ghi chú: X — người dùng có quyền thực hiện chức năng. X* — chỉ có quyền trên bản ghi thuộc về mình.'));
children.push(table(
  ['Phân hệ / Chức năng', 'Nhân viên', 'Khách hàng', 'Hệ thống'],
  [
    ['Phân hệ 1: Chung', '', '', ''],
    ['Đăng nhập', 'X', 'X', ''],
    ['Đăng xuất', 'X', 'X', ''],
    ['Phân hệ 2: Quản lý tài khoản', '', '', ''],
    ['Đăng ký tài khoản (OTP email)', '', 'X', ''],
    ['Cập nhật hồ sơ', '', 'X*', ''],
    ['Đổi mật khẩu', '', 'X*', ''],
    ['Quản lý tài khoản người dùng', 'X', '', ''],
    ['Phân hệ 3: Quản lý phương tiện', '', '', ''],
    ['Xem danh sách xe', '', 'X*', ''],
    ['Đăng ký xe (OTP)', '', 'X*', ''],
    ['Xóa xe', '', 'X*', ''],
    ['Phân hệ 4: Quản lý dịch vụ & đặt lịch', '', '', ''],
    ['Xem dịch vụ', '', 'X', ''],
    ['Quản lý dịch vụ', 'X', '', ''],
    ['Tạo đặt lịch', '', 'X*', ''],
    ['Xem đặt lịch', '', 'X*', ''],
    ['Xác nhận đặt lịch', 'X', '', ''],
    ['Hủy đặt lịch', 'X', '', ''],
    ['Phân hệ 5: Quản lý hàng chờ', '', '', ''],
    ['Xem hàng chờ hôm nay', 'X', '', ''],
    ['Check-in khách hàng', 'X', '', ''],
    ['Quét biển số (LPR)', 'X', '', ''],
    ['Chuyển trạng thái rửa xe', 'X', '', ''],
    ['Cập nhật hàng chờ', 'X', '', ''],
    ['Check-out khách hàng', 'X', '', ''],
    ['Hủy lượt hàng chờ', 'X', '', ''],
    ['Phân hệ 6: Thông báo', '', '', ''],
    ['Xem thông báo', '', 'X*', ''],
    ['Đánh dấu thông báo đã đọc', '', 'X*', ''],
    ['Gửi email OTP / xác nhận / hoàn tất', '', '', 'X'],
  ],
  [4526, 1500, 1500, 1500],
));
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- 3. Đặc tả Use Case ---------------- */
children.push(h1('3. Đặc tả Use Case'));
children.push(p('Phần này mô tả chi tiết yêu cầu chức năng của hệ thống: tương tác của tác nhân, điều kiện trước/sau, luồng hoạt động và quy tắc nghiệp vụ.'));

children.push(h2('3.1 Chung'));
uc1.general.forEach((u) => children.push(...renderUC(u, brCounter)));
children.push(h2('3.2 Quản lý tài khoản'));
uc1.account.forEach((u) => children.push(...renderUC(u, brCounter)));
children.push(h2('3.3 Quản lý phương tiện'));
uc1.vehicle.forEach((u) => children.push(...renderUC(u, brCounter)));
children.push(h2('3.4 Quản lý dịch vụ & đặt lịch'));
uc2.booking.forEach((u) => children.push(...renderUC(u, brCounter)));
children.push(h2('3.5 Quản lý hàng chờ'));
uc2.queue.forEach((u) => children.push(...renderUC(u, brCounter)));
children.push(h2('3.6 Thông báo'));
uc2.notification.forEach((u) => children.push(...renderUC(u, brCounter)));

children.push(h2('3.7 Quy tắc nghiệp vụ chung'));
children.push(table(
  ['Mã CBR', 'Tên quy tắc', 'Mô tả'],
  [
    ['CBR1', 'Bắt buộc xác thực', 'Mọi endpoint trừ Đăng nhập, Đăng ký và Gửi OTP đều yêu cầu session hợp lệ phía máy chủ. Yêu cầu không hợp lệ trả về HTTP 401.'],
    ['CBR2', 'Kiểm soát vai trò', 'Endpoint dành cho Nhân viên từ chối session của Khách hàng. Yêu cầu vi phạm trả về HTTP 403.'],
    ['CBR3', 'Vòng đời OTP', 'Mọi OTP gồm 6 chữ số, sinh ngẫu nhiên, lưu dạng băm, hết hạn sau 5 phút và chỉ dùng một lần. Sau khi dùng, IsUsed = true.'],
    ['CBR4', 'Chuẩn hóa biển số', 'Mọi biển số được chuẩn hóa trước khi lưu và tra cứu: chuyển chữ hoa, bỏ khoảng trắng và dấu gạch (VD: "51a-123.45" → "51A12345"). Áp dụng cho cả biển số nhập tay và kết quả từ LPR.'],
    ['CBR5', 'Kiểm tra quyền sở hữu', 'Trước mọi thao tác trên tài nguyên (Vehicle, Booking, Notification), hệ thống kiểm tra tài nguyên thuộc về khách hàng đang đăng nhập.'],
    ['CBR6', 'Tính giá', 'FinalPrice = tổng PriceSnapshot của các dịch vụ đã chọn. Giá lưu dưới dạng số nguyên VND.'],
    ['CBR7', 'Giới hạn công suất', 'Tối đa 3 lượt đặt trên mỗi khung 1 giờ. Xe vãng lai (walk-in) không chịu giới hạn công suất đặt lịch.'],
    ['CBR8', 'Thứ tự hàng chờ FIFO', 'Mục hàng chờ được sắp xếp theo CheckInAt tăng dần — đến trước phục vụ trước.'],
    ['CBR9', 'Toàn vẹn dữ liệu', 'Mọi thao tác nhiều bước (tạo booking, check-out) được bọc trong giao dịch cơ sở dữ liệu để tránh ghi dữ liệu dở dang.'],
  ],
  [1100, 2200, 5726],
));
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- 4. Mockup màn hình ---------------- */
children.push(h1('4. Màn hình Mockup'));
children.push(h2('4.1 Sơ đồ luồng màn hình (Dialog Map)'));
children.push(p('Sơ đồ luồng màn hình (dialog map) thể hiện toàn bộ màn hình và cách điều hướng giữa chúng: mỗi mũi tên là một hành động chuyển màn (nhấn nút, mở/đóng hộp thoại). Màn hình trung tâm (hub) tô đậm; hộp thoại vẽ nét đứt.', { run: { italics: true } }));
const dmDir = path.join(__dirname, '..', '..', 'diagrams', 'diagram picture', 'DialogMap');
[['4.1.1 Luồng màn hình — Khách hàng (Mobile Web)', 'DialogMap-1. Khách hàng (Mobile Web).drawio.png', 520, 224],
 ['4.1.2 Luồng màn hình — Nhân viên (Desktop)', 'DialogMap-2. Nhân viên (Desktop).drawio.png', 540, 145]].forEach(function (m) {
  children.push(h3(m[0]));
  const dfp = path.join(dmDir, m[1]);
  if (fs.existsSync(dfp)) {
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
      children: [ new ImageRun({ type: 'png', data: fs.readFileSync(dfp), transformation: { width: m[2], height: m[3] } }) ] }));
    children.push(p('Nguồn: diagrams/DialogMap.drawio (' + (m[1].indexOf('Customer') >= 0 ? 'sheet 1 — Khách hàng' : 'sheet 2 — Nhân viên') + ').', { run: { italics: true, size: 18 } }));
  } else {
    children.push(p('[Dialog map — xem diagrams/DialogMap.drawio]', { run: { italics: true } }));
  }
});
const mockups = [
  ['4.2 Chung', ['4.2.1 Màn hình Đăng nhập', '4.2.2 Trang chủ (Landing Page)']],
  ['4.3 Quản lý tài khoản & phương tiện', ['4.3.1 Màn hình Đăng ký', '4.3.2 Đăng ký — Xác thực OTP', '4.3.3 Cài đặt hồ sơ — Xem & chỉnh sửa', '4.3.4 Màn hình Đổi mật khẩu', '4.3.5 Xe của tôi — Danh sách', '4.3.6 Xe của tôi — Thêm xe (OTP)', '4.3.7 Quản lý tài khoản người dùng (Nhân viên)']],
  ['4.4 Quản lý dịch vụ & đặt lịch', ['4.4.2 Màn hình Danh sách dịch vụ', '4.4.3 Quản lý dịch vụ (Nhân viên)', '4.4.3 Tạo đặt lịch — Chọn dịch vụ & ngày giờ', '4.4.4 Tạo đặt lịch — Tóm tắt & xác nhận', '4.4.5 Đặt lịch của tôi — Danh sách & theo dõi hàng chờ', '4.4.6 Danh sách booking chờ xác nhận (Nhân viên)']],
  ['4.5 Quản lý hàng chờ', ['4.5.1 Bảng hàng chờ — Danh sách hôm nay', '4.5.2 Bảng hàng chờ — Chuyển trạng thái', '4.5.3 Bảng hàng chờ — Cập nhật / ghi chú', '4.5.4 Bảng hàng chờ — Hộp thoại Check-in (kèm nút Quét biển số)', '4.5.6 Bảng hàng chờ — Hộp thoại Check-out', '4.5.6 Bảng hàng chờ — Xác nhận hủy lượt']],
  ['4.6 Thông báo', ['4.6.1 Thông báo — Danh sách', '4.6.2 Thông báo — Chi tiết / đánh dấu đã đọc']],
];
const mockDir = path.join(__dirname, '..', '..', 'diagrams', 'diagram picture', 'mockups');
const P = [232, 390];   // phone size
const D = [478, 326];   // desktop size
const mockMap = {
  '4.2.1 Màn hình Đăng nhập': ['login.png', ...P],
  '4.2.2 Trang chủ (Landing Page)': ['landing.png', ...P],
  '4.3.1 Màn hình Đăng ký': ['register.png', ...P],
  '4.3.2 Đăng ký — Xác thực OTP': ['register_otp.png', ...P],
  '4.3.3 Cài đặt hồ sơ — Xem & chỉnh sửa': ['profile.png', ...P],
  '4.3.4 Màn hình Đổi mật khẩu': ['change_password.png', ...P],
  '4.3.5 Xe của tôi — Danh sách': ['my_vehicles.png', ...P],
  '4.3.6 Xe của tôi — Thêm xe (OTP)': ['add_vehicle_otp.png', ...P],
  '4.3.7 Quản lý tài khoản người dùng (Nhân viên)': ['manage_users.png', ...D],
  '4.4.2 Màn hình Danh sách dịch vụ': ['services_list.png', ...P],
  '4.4.3 Quản lý dịch vụ (Nhân viên)': ['manage_services.png', ...D],
  '4.4.3 Tạo đặt lịch — Chọn dịch vụ & ngày giờ': ['create_booking.png', ...P],
  '4.4.4 Tạo đặt lịch — Tóm tắt & xác nhận': ['booking_summary.png', ...P],
  '4.4.5 Đặt lịch của tôi — Danh sách & theo dõi hàng chờ': ['my_bookings.png', ...P],
  '4.4.6 Danh sách booking chờ xác nhận (Nhân viên)': ['pending_bookings.png', ...D],
  '4.5.1 Bảng hàng chờ — Danh sách hôm nay': ['staff_queue.png', ...D],
  '4.5.2 Bảng hàng chờ — Chuyển trạng thái': ['advance_status.png', ...D],
  '4.5.3 Bảng hàng chờ — Cập nhật / ghi chú': ['update_queue.png', ...D],
  '4.5.4 Bảng hàng chờ — Hộp thoại Check-in (kèm nút Quét biển số)': ['checkin_dialog.png', ...D],
  '4.5.6 Bảng hàng chờ — Hộp thoại Check-out': ['checkout_dialog.png', ...D],
  '4.5.6 Bảng hàng chờ — Xác nhận hủy lượt': ['cancel_queue.png', ...D],
  '4.6.1 Thông báo — Danh sách': ['notifications.png', ...P],
  '4.6.2 Thông báo — Chi tiết / đánh dấu đã đọc': ['notification_detail.png', ...P],
};
mockups.forEach(([sec, items]) => {
  children.push(h2(sec));
  items.forEach((it) => {
    children.push(h3(it));
    const m = mockMap[it];
    const fp = m ? path.join(mockDir, m[0]) : null;
    if (fp && fs.existsSync(fp)) {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 80 },
        children: [ new ImageRun({ type: 'png', data: fs.readFileSync(fp), transformation: { width: m[1], height: m[2] } }) ],
      }));
      children.push(p('Mockup minh họa — nguồn: diagrams/mockups/' + m[0] + ' (bản đầy đủ tại diagrams/Mockups.html).', { run: { italics: true, size: 18 } }));
    } else {
      children.push(p('[Mockup màn hình — bản wireframe xem tại diagrams/Mockups.html; các màn còn lại sẽ bổ sung]', { run: { italics: true } }));
    }
  });
});
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- 5. Yêu cầu khác ---------------- */
children.push(h1('5. Yêu cầu khác'));
children.push(h2('5.1 Yêu cầu hiệu năng'));
children.push(table(
  ['Yêu cầu', 'Chỉ tiêu'],
  [
    ['Thời gian phản hồi API (truy vấn tiêu chuẩn: danh sách booking, hàng chờ)', '< 2 giây ở bách phân vị 95'],
    ['Gửi email OTP', '< 60 giây kể từ khi yêu cầu'],
    ['Tạo booking (bao gồm kiểm tra hợp lệ và ghi CSDL)', '< 3 giây'],
    ['Thời gian trả kết quả nhận dạng biển số (LPR) từ khi chụp ảnh', '< 5 giây'],
    ['Số người dùng đồng thời (mục tiêu giai đoạn đầu)', '50 phiên đồng thời'],
  ],
  [6326, 2700],
));
children.push(h2('5.2 Yêu cầu bảo mật'));
children.push(table(
  ['Yêu cầu', 'Mô tả'],
  [
    ['Lưu trữ mật khẩu', 'Mọi mật khẩu lưu dưới dạng băm mật mã (bcrypt hoặc tương đương). Không bao giờ lưu hoặc ghi log mật khẩu thuần.'],
    ['Quản lý phiên', 'Session phía máy chủ, idle timeout 8 giờ. Cookie đặt cờ HttpOnly và Secure trong môi trường production.'],
    ['Kiểm soát xác thực', 'Mọi endpoint không công khai kiểm tra session phía máy chủ. Session thiếu/không hợp lệ trả về HTTP 401.'],
    ['Kiểm soát phân quyền', 'Kiểm soát truy cập theo vai trò. Endpoint Nhân viên từ chối session Khách hàng (HTTP 403).'],
    ['Bảo mật OTP', 'OTP 6 chữ số, lưu dạng băm, hiệu lực 5 phút, chỉ dùng một lần (CBR3).'],
    ['Lưu trữ thông tin bí mật', 'Thông tin đăng nhập CSDL, mật khẩu SMTP và khóa API của 3rd-party LPR API lưu trong biến môi trường, không lưu trong mã nguồn.'],
    ['Kiểm tra dữ liệu vào', 'Mọi dữ liệu người dùng nhập được kiểm tra phía máy chủ. Biển số kiểm tra bằng regex. Ngày đặt lịch kiểm tra theo quy tắc nghiệp vụ.'],
    ['HTTPS', 'Toàn bộ lưu lượng production phải dùng HTTPS. HTTP được chuyển hướng sang HTTPS.'],
  ],
  [2400, 6626],
));
children.push(h2('5.3 Yêu cầu khả dụng (Usability)'));
children.push(table(
  ['Yêu cầu', 'Mô tả'],
  [
    ['Thiết kế responsive', 'Kênh Khách hàng là Mobile Web: hệ thống phải hoạt động đầy đủ trên trình duyệt di động (từ 375px); màn hình Nhân viên hỗ trợ desktop (từ 1024px).'],
    ['Thông điệp lỗi', 'Mọi thông điệp hệ thống (MSG1–MSG25) phải dễ hiểu, cụ thể và có tính hướng dẫn.'],
    ['Trạng thái đang tải', 'Mọi thao tác bất đồng bộ phải hiển thị chỉ báo đang tải để tránh gửi trùng.'],
    ['Xác nhận đặt lịch', 'Khách hàng phải được xem bản tóm tắt trước khi xác nhận đặt lịch cuối cùng.'],
  ],
  [2400, 6626],
));
children.push(h2('5.4 Yêu cầu sẵn sàng & tin cậy'));
children.push(table(
  ['Yêu cầu', 'Chỉ tiêu'],
  [
    ['Thời gian hoạt động trong giờ làm việc (07:00–22:00 ICT)', '≥ 99%'],
    ['Cửa sổ bảo trì theo kế hoạch', 'Chỉ ngoài giờ làm việc'],
    ['Thao tác ghi CSDL', 'Mọi thao tác ghi nhiều bước bọc trong giao dịch CSDL (CBR9)'],
    ['Chống tranh chấp dữ liệu', 'Giới hạn công suất khung giờ được kiểm tra trong giao dịch để tránh đặt vượt số lượng'],
    ['Dự phòng khi LPR lỗi', 'Khi Camera LPR hoặc 3rd-party LPR API không khả dụng, nhân viên luôn có thể nhập biển số thủ công (UC17)'],
  ],
  [5326, 3700],
));
children.push(h2('5.5 Yêu cầu mở rộng (Scalability)'));
children.push(p('Hệ thống xây dựng trên Supabase PostgreSQL, hỗ trợ mở rộng. Backend ASP.NET Core không lưu trạng thái phía ứng dụng (session lưu phía máy chủ), cho phép triển khai cân bằng tải trong tương lai. Hệ thống được thiết kế ban đầu cho một trung tâm rửa xe; hỗ trợ đa chi nhánh nằm ngoài phạm vi phiên bản này.'));
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- 6. Tích hợp ---------------- */
children.push(h1('6. Tích hợp'));
children.push(table(
  ['Tích hợp', 'Nhà cung cấp / Công nghệ', 'Mục đích', 'Cấu hình'],
  [
    ['Cơ sở dữ liệu quan hệ', 'PostgreSQL qua Supabase', 'Kho dữ liệu chính cho mọi thực thể. Truy cập qua EF Core (Npgsql).', 'Chuỗi kết nối trong biến môi trường DefaultConnection. Bắt buộc SSL.'],
    ['Email (OTP & thông báo)', 'Gmail SMTP qua thư viện MailKit', 'Gửi mã OTP (đăng ký, đăng ký xe, đổi mật khẩu), email xác nhận đặt lịch và email thông báo rửa xe hoàn tất.', 'SMTP Host: smtp.gmail.com, Port: 587 (STARTTLS). Thông tin đăng nhập trong cấu hình Smtp:Username, Smtp:Password.'],
    ['Nhận dạng biển số (LPR)', 'Camera LPR + 3rd-party LPR API (API ngoài)', 'Camera chụp ảnh biển số; 3rd-party LPR API trích xuất chuỗi biển số từ ảnh và trả kết quả về LPR Service của hệ thống.', 'Endpoint và khóa API của 3rd-party LPR API lưu trong biến môi trường. Timeout 5 giây; khi lỗi, chuyển sang nhập thủ công.'],
    ['Frontend Mobile Web', 'React 18 + Vite', 'Ứng dụng web một trang (SPA) tối ưu cho di động, giao tiếp với backend qua JSON REST API.', 'Môi trường phát triển: http://localhost:5173. Chính sách CORS cho phép credentials từ origin này.'],
    ['Phục vụ tệp tĩnh', 'ASP.NET Core static files middleware', 'Phục vụ bản build của SPA từ thư mục wwwroot trong production.', 'Cấu hình fallback route trả về index.html để hỗ trợ React Router.'],
  ],
  [1700, 1900, 2800, 2626],
));
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- 7. Di trú dữ liệu ---------------- */
children.push(h1('7. Di trú dữ liệu'));
children.push(h2('7.1 Phạm vi di trú'));
children.push(p('Auto-Wash là hệ thống xây mới (greenfield); không có dữ liệu rửa xe cũ. Các dữ liệu sau sẽ được khởi tạo (seed) khi triển khai lần đầu:'));
children.push(table(
  ['Nhóm dữ liệu', 'Nguồn', 'Hành động'],
  [
    ['Danh mục dịch vụ (Rửa cơ bản, Rửa cao cấp, Rửa đặc biệt, các add-on)', 'Nhân viên nhập thủ công', 'Nhập qua màn hình Quản lý dịch vụ (UC11) hoặc script SQL seed'],
    ['Tài khoản Nhân viên', 'Bộ phận IT', 'Chèn trực tiếp với mật khẩu đã băm (không cần OTP cho nhân viên)'],
    ['Tài khoản Khách hàng ban đầu (nếu chuyển từ hệ thống trước)', 'Xuất từ hệ thống trước', 'Chuyển đổi và nhập qua script di trú'],
  ],
  [4026, 2200, 3000],
));
children.push(h2('7.2 Ánh xạ dữ liệu'));
children.push(table(
  ['Trường hệ thống cũ', 'Thực thể.Trường Auto-Wash', 'Quy tắc chuyển đổi'],
  [
    ['Số điện thoại khách hàng', 'Account.Phone', 'Chuẩn hóa về định dạng 10 chữ số; bỏ mã quốc gia'],
    ['Tên khách hàng', 'Account.FullName', 'Cắt khoảng trắng thừa; viết hoa chữ cái đầu'],
    ['Biển số xe', 'Vehicle.LicensePlate', 'Chuẩn hóa: chữ hoa, bỏ khoảng trắng/gạch ngang (CBR4)'],
    ['Bản ghi dịch vụ lịch sử', 'Booking (Status = Hoàn tất)', 'Ánh xạ tên dịch vụ sang ServiceId; lấy FinalPrice từ bản ghi gốc'],
    ['Tổng chi tiêu tích lũy', 'Customer.TotalSpend', 'Nhập tổng chi tiêu lịch sử'],
  ],
  [2600, 2800, 3626],
));
children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------------- 8. Phụ lục ---------------- */
children.push(h1('8. Phụ lục'));
children.push(h2('8.1 Danh sách thông điệp'));
children.push(table(
  ['Mã MSG', 'Nội dung thông điệp', 'Nguồn kích hoạt'],
  [
    ['MSG1', 'Vui lòng nhập đầy đủ các trường bắt buộc.', 'UC1 — Đăng nhập: thiếu thông tin'],
    ['MSG2', 'Tên đăng nhập hoặc mật khẩu không đúng.', 'UC1 — Đăng nhập: sai thông tin'],
    ['MSG3', 'Email đã được đăng ký. Vui lòng dùng email khác.', 'UC3 — Đăng ký: trùng email'],
    ['MSG4', 'Mã OTP đã hết hạn. Vui lòng yêu cầu mã mới.', 'UC3/UC5/UC8 — OTP hết hạn'],
    ['MSG5', 'Mã OTP không hợp lệ hoặc đã được sử dụng.', 'Mọi lỗi xác thực OTP'],
    ['MSG6', 'Không thể xóa xe đã có lịch sử đặt lịch.', 'UC9 — Xóa xe có booking'],
    ['MSG7', 'Xe này đã được đăng ký trong tài khoản của bạn.', 'UC8 — Trùng biển số'],
    ['MSG8', 'Biển số không đúng định dạng Việt Nam (VD: 51A-12345).', 'UC8 — Biển số không hợp lệ'],
    ['MSG9', 'Xe được chọn không thuộc tài khoản của bạn.', 'UC12 — Sai quyền sở hữu xe'],
    ['MSG10', 'Thời gian đặt lịch phải cách hiện tại ít nhất 15 phút.', 'UC12 — Thời gian quá gần'],
    ['MSG11', 'Khung giờ này đã đầy (tối đa 3 lượt đặt/giờ). Vui lòng chọn giờ khác.', 'UC12 — Vượt công suất'],
    ['MSG12', 'Bạn đã có lịch đặt cho xe này trong khung giờ đã chọn.', 'UC12 — Trùng booking'],
    ['MSG13', 'Dịch vụ được chọn hiện không khả dụng.', 'UC12 — Dịch vụ bị vô hiệu hóa'],
    ['MSG14', 'Đặt lịch thành công. Thông báo xác nhận đã được gửi.', 'UC12 — Lưu booking thành công'],
    ['MSG15', 'Cập nhật hồ sơ thành công.', 'UC4 — Cập nhật hồ sơ thành công'],
    ['MSG16', 'Đổi mật khẩu thành công.', 'UC5 — Đổi mật khẩu thành công'],
    ['MSG17', 'Đăng ký xe thành công.', 'UC8 — Lưu xe thành công'],
    ['MSG18', 'Xóa xe thành công.', 'UC9 — Xóa xe thành công'],
    ['MSG19', 'Không thể hủy một lượt rửa đã hoàn tất.', 'UC22 — Hủy lượt đã hoàn tất'],
    ['MSG20', 'Số điện thoại đã được sử dụng bởi tài khoản khác.', 'UC3/UC4 — Trùng số điện thoại'],
    ['MSG21', 'Cập nhật dịch vụ thành công.', 'UC11 — Lưu dịch vụ thành công'],
    ['MSG22', 'Cập nhật tài khoản thành công.', 'UC6 — Cập nhật tài khoản thành công'],
    ['MSG23', 'Không nhận diện được biển số. Vui lòng nhập thủ công.', 'UC18 — LPR thất bại'],
    ['MSG24', 'Xác nhận đặt lịch thành công.', 'UC14 — Xác nhận booking'],
    ['MSG25', 'Hủy đặt lịch thành công.', 'UC15 — Hủy booking'],
    ['MSG26', 'Tài khoản của bạn đã bị vô hiệu hóa. Vui lòng liên hệ nhân viên để được hỗ trợ.', 'UC1 — Đăng nhập: tài khoản bị khóa (IsActive = false)'],
    ['MSG27', 'Bạn không thể tự vô hiệu hóa tài khoản đang đăng nhập.', 'UC6 — Quản lý tài khoản: chặn tự khóa tài khoản mình'],
    ['MSG28', 'Bạn chưa đăng ký xe nào. Vui lòng thêm xe trước khi đặt lịch.', 'UC12 — Tạo đặt lịch: chưa có xe'],
    ['MSG29', 'Booking này vừa được xử lý bởi thao tác khác. Vui lòng tải lại danh sách.', 'UC14 — Xác nhận đặt lịch: tranh chấp trạng thái đồng thời'],
    ['MSG30', 'Không thể hủy booking ở trạng thái hiện tại.', 'UC15 — Hủy đặt lịch: sai trạng thái'],
    ['MSG31', 'Không thể chuyển trạng thái: lượt rửa đã hoàn tất hoặc đã hủy, hoặc không đúng thứ tự công đoạn.', 'UC19 — Chuyển trạng thái rửa xe: chuyển trạng thái không hợp lệ'],
    ['MSG32', 'Giá và thời lượng dịch vụ phải là số hợp lệ (giá ≥ 0, thời lượng ≥ 1 phút).', 'UC11 — Quản lý dịch vụ: dữ liệu không hợp lệ'],
  ],
  [1400, 4626, 3000],
));
children.push(spacer());

/* ---------------- 8.2 Mẫu email ---------------- */
children.push(h2('8.2 Mẫu email'));

children.push(h3('8.2.1 ET 1: Mã xác thực OTP'));
children.push(table(
  ['Trường', 'Giá trị'],
  [
    ['Kích hoạt', 'Yêu cầu OTP cho: đăng ký tài khoản (UC3), đổi mật khẩu (UC5), đăng ký xe (UC8). Đăng nhập (UC1) không dùng OTP.'],
    ['Người nhận', 'Địa chỉ email thực hiện yêu cầu OTP.'],
    ['Tiêu đề', 'Mã xác thực Auto-Wash của bạn'],
    ['Nội dung', 'Mã xác thực của bạn là: [OTP_CODE]. Mã có hiệu lực trong 5 phút. Không chia sẻ mã này với bất kỳ ai.'],
  ],
  [2000, 7026],
));
children.push(spacer());

children.push(h3('8.2.2 ET 2: Xác nhận đặt lịch'));
children.push(table(
  ['Trường', 'Giá trị'],
  [
    ['Kích hoạt', 'Khách hàng tạo đặt lịch thành công (UC12).'],
    ['Người nhận', 'Khách hàng đã đặt lịch.'],
    ['Tiêu đề', 'Xác nhận đặt lịch — Auto-Wash'],
    ['Nội dung', 'Chào [FullName], lịch đặt của bạn đã được ghi nhận. Chi tiết: Xe: [LicensePlate] | Dịch vụ: [ServiceName] | Ngày giờ: [ScheduledAt] | Thành tiền: [FinalPrice] VND. Hẹn gặp bạn!'],
  ],
  [2000, 7026],
));
children.push(spacer());

children.push(h3('8.2.3 ET 3: Rửa xe hoàn tất'));
children.push(table(
  ['Trường', 'Giá trị'],
  [
    ['Kích hoạt', 'Nhân viên hoàn tất check-out (UC21).'],
    ['Người nhận', 'Khách hàng có xe vừa được rửa.'],
    ['Tiêu đề', 'Xe của bạn đã rửa xong — Auto-Wash'],
    ['Nội dung', 'Chào [FullName], xe [LicensePlate] của bạn đã được rửa xong. Cảm ơn bạn đã lựa chọn Auto-Wash!'],
  ],
  [2000, 7026],
));
children.push(spacer());
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240 }, children: [t('--- Hết tài liệu ---', { italics: true })] }));

/* ================= Lắp ráp tài liệu & xuất ================= */
const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 360, hanging: 180 } } } },
        ],
      },
    ],
  },
  styles: {
    default: { document: { run: { font: 'Times New Roman', size: 24 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: 'Arial', size: 30, bold: true, color: '1F3864' },
        paragraph: { spacing: { before: 280, after: 140 }, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: 'Arial', size: 26, bold: true, color: '2E5496' },
        paragraph: { spacing: { before: 200, after: 100 }, keepNext: true } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: 'Arial', size: 24, bold: true, color: '1F4E79' },
        paragraph: { spacing: { before: 160, after: 80 }, keepNext: true } },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              tabStops: [{ type: TabStopType.RIGHT, position: 9026 }],
              border: { bottom: { color: 'AAAAAA', space: 4, style: 'single', size: 4 } },
              children: [ t('Auto-Wash — SRS v2.8', { size: 18, color: '666666' }),
                          new TextRun({ text: '\tFPT University — SWP391', size: 18, color: '666666' }) ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [ t('Trang ', { size: 18, color: '666666' }),
                          new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '666666' }),
                          t(' / ', { size: 18, color: '666666' }),
                          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: '666666' }) ],
            }),
          ],
        }),
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  const outName = 'Auto-Wash_SRS_v2.8.docx';
  fs.writeFileSync(outName, buf);
  console.log('Da tao', outName, '(' + buf.length + ' bytes)');
});

