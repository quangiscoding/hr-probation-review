## 2. Quy tắc lập trình bắt buộc (Engineering Guidelines)

### A. Phân tách Feature Folder (`src/<feature-name>/`)

- Mọi logic nghiệp vụ, hàm bổ trợ (helper/util), cấu hình hằng số (config) và giao diện người dùng (HTML dialog/sidebar/template) thuộc về một nghiệp vụ cụ thể phải được đóng gói tập trung trong một thư mục duy nhất: `src/<feature-name>/`.
- Tuyệt đối không khai báo các biến, hàm tiện ích hoặc logic dùng riêng cho tính năng này ở thư mục gốc của `src/` nhằm đảm bảo tính cô lập (encapsulation) và tránh làm phân mảnh codebase.

### B. Quy tắc Service Pattern (`<FeatureName>Service`)

- Toàn bộ nghiệp vụ chính của tính năng (truy vấn dữ liệu bảng tính, xử lý logic lõi, tương tác với Google Drive/Gmail/Calendar/External API, khóa dòng chống lặp) phải được đóng gói trọn vẹn trong một class duy nhất: `<FeatureName>Service`.
- Tất cả các phương thức trong Service được thiết kế dưới dạng **`static methods`** để truy xuất trực tiếp mà không cần khởi tạo instance.
- **Contract trả về (Return Contract):**
  - Mọi phương thức trong `<FeatureName>Service` chỉ trả về **dữ liệu thuần** (Plain Object / JSON) phản ánh kết quả thực thi (ví dụ: `{ ok: true, message: string, summary: Object, ... }`) hoặc ném ngoại lệ có định danh (`throw new WorkflowError(...)`).
  - **Tuyệt đối không** gọi trực tiếp các phương thức UI của Google Workspace (`SpreadsheetApp.getUi()`, `alert()`, `showModalDialog()`, `showSidebar()`) bên trong tầng Service.

### C. Quy tắc Utility Pattern (`<FeatureName>Util`)

- Mọi hàm hỗ trợ xử lý dữ liệu thuần túy (chuẩn hóa chuỗi header, định dạng ngày tháng, tính toán khoảng thời gian, regex trích xuất ID Google Drive/Folder, kiểm tra định dạng email) phải được bọc trong class: `<FeatureName>Util`.
- Không khai báo các hàm helper tự do ngoài global scope để ngăn chặn triệt để nguy cơ xung đột định danh (name collision) giữa các module trong môi trường runtime phẳng của Google Apps Script.
- Hàm trích xuất Google Drive / Folder ID phải tuân thủ chuẩn an toàn nghiêm ngặt:
  - Khớp theo các cấu trúc URL chuẩn của Google: sau `/d/`, `/folders/` hoặc tham số `?id=`.
  - Giới hạn độ dài ID chuẩn từ 25 đến 50 ký tự và có ký tự biên chặn (`[/?&#]|$`) để loại bỏ hoàn toàn các chuỗi token, hash hay tham số URL thừa.
  - Fallback an toàn về chuỗi ID thô nếu người dùng dán trực tiếp mã ID.

### D. File `main.js` chỉ làm nhiệm vụ Controller

- File `main.js` giữ vai trò là **Controller / Entry Point** duy nhất của toàn bộ hệ thống.
- Chỉ chứa các trigger hệ thống (`onOpen`, `onEdit`, time-driven triggers) và các hàm điều phối menu:
  - Tiếp nhận sự kiện người dùng kích hoạt từ Custom Menu.
  - Gọi các thông báo xác nhận hành động (Confirm Dialog) nếu là thao tác nhạy cảm.
  - Điều phối và ủy quyền thực thi nghiệp vụ hoàn toàn cho `<FeatureName>Service`.
  - Nhận kết quả dữ liệu thuần từ Service và hiển thị giao diện tương ứng (Toast, Alert, Modal Dialog, hoặc Sidebar) qua module UI feedback chung.
- Tuyệt đối không viết câu lệnh truy vấn dữ liệu Sheet/Drive hoặc xử lý thuật toán nghiệp vụ trực tiếp trong `main.js`.

### E. Tối ưu hóa hiệu năng Google Apps Script (Batch Reading & RAM Cache)

- **Triệt tiêu các lời gọi API lặp (Zero O(N) Sheet API calls):**
  - Đọc toàn bộ dữ liệu bảng tính cần thiết về bộ nhớ RAM một lần duy nhất qua phương thức `.getValues()`.
  - Sử dụng hàm ánh xạ tiêu đề động (`getHeaderColumnMap`) để xác định chỉ số cột (1-indexed) theo dòng header thực tế, tuyệt đối không gán cứng chỉ số cột dạng số.
  - Tạo cấu trúc `Map` hoặc `Set` trên bộ nhớ RAM dựa trên khóa định danh duy nhất (ví dụ: `Email`, `Alloc Code`, hoặc `ID`) để thực hiện tra cứu và đối soát chéo liên Sheet đạt độ phức tạp O(1).
- **Cơ chế chống chạy lặp (Anti-duplication Mechanism):**
  - Luôn kiểm tra trạng thái đã xử lý (thông qua cột lưu vết timestamp hoặc Developer Metadata gắn trên dòng) trước khi thực thi gửi email, tạo draft hay ghi nhận dữ liệu mới, đảm bảo tính toàn vẹn và không trùng lặp tác vụ.

---

## 3. Quy chuẩn bảng dữ liệu & Mapping liên Sheet

Khi một tính năng cần phối hợp và đối soát dữ liệu giữa nhiều bảng tính (hoặc nhiều tab Sheet), cấu trúc ánh xạ phải được đặc tả rõ ràng theo mẫu chuẩn sau:

| Bảng tính / Vai trò                                                 | Vị trí Sheet / Tab  | Khóa định danh (Matching Key) | Tên cột nghiệp vụ           | Kiểu dữ liệu    | Mục đích / Xử lý nghiệp vụ                                        |
| :------------------------------------------------------------------ | :------------------ | :---------------------------- | :-------------------------- | :-------------- | :---------------------------------------------------------------- |
| **`<Trigger / Control Sheet>`**<br>_(Bảng điều khiển & quét luồng)_ | Tên tab / Sheet GID | `<Primary_Key>`               | `<Cột thời gian / Trigger>` | Ngày / Giờ      | Mốc thời gian kích hoạt luồng tự động hoặc điều kiện lọc dữ liệu. |
|                                                                     |                     |                               | `<Cột đối soát chính>`      | Email / Mã ID   | Khóa chính map 1-1 với bảng dữ liệu nguồn.                        |
|                                                                     |                     |                               | `<Cột người phụ trách>`     | Email / Text    | Thông tin người thực hiện / nhận thông báo liên quan (CC/BCC).    |
|                                                                     |                     |                               | `<Cột trạng thái / Log>`    | Ngày giờ / Text | Lưu vết thời gian hoặc trạng thái hoàn tất để chống chạy lặp.     |
| **`<Source / Master Data Sheet>`**<br>_(Bảng lưu trữ hồ sơ gốc)_    | Tên tab / Sheet GID | `<Foreign_Key>`               | `<Cột đối soát tương ứng>`  | Email / Mã ID   | Khóa khớp nối với `<Primary_Key>` từ Bảng điều khiển.             |
|                                                                     |                     |                               | `<Cột mốc thời gian phụ>`   | Ngày / Giờ      | Mốc đối soát bổ sung (ví dụ: tính lùi/tiến số ngày quy định).     |
|                                                                     |                     |                               | `<Cột tài nguyên Drive>`    | URL Drive       | Đường dẫn thư mục hoặc file tài liệu cần tải/đính kèm.            |
