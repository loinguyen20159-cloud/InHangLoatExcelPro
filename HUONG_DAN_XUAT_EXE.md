# IN HÀNG LOẠT EXCEL PRO (`InHangLoatExcelPro.exe`)

Ứng dụng desktop Windows hỗ trợ in hàng loạt file Excel (`XLSX`, `XLSM`, `XLS`, `XLSB`) tự động chọn sheet `WL`, giữ nguyên vùng in & tỷ lệ gốc.

## Vì sao mình xuất cho bạn gói mã nguồn + cấu hình build tự động?
Môi trường đám mây hiện tại chạy trên Linux nên không thể trực tiếp biên dịch file `.exe` của Windows. Tuy nhiên, mình đã đóng gói sẵn toàn bộ mã nguồn hoàn chỉnh kèm **2 cách cực kỳ đơn giản** để bạn có file `InHangLoatExcelPro.exe` chạy trực tiếp mà **máy bạn hoàn toàn không cần cài Python**:

### Cách 1: Nhờ GitHub build tự động miễn phí (Máy bạn KHÔNG cần cài bất cứ thứ gì)
1. Tạo một tài khoản miễn phí tại [github.com](https://github.com) và tạo 1 Repository mới (ví dụ `InHangLoatExcelPro`).
2. Giải nén file ZIP này và tải toàn bộ các file bên trong lên Repository đó (giữ nguyên thư mục `.github/workflows/build_exe.yml`).
3. Vào tab **Actions** trên GitHub -> chọn workflow **Build InHangLoatExcelPro.exe** -> bấm **Run workflow**.
4. Chờ khoảng 1–2 phút, tải file **`InHangLoatExcelPro-Windows.zip`** ở mục **Artifacts** về giải nén là có ngay `InHangLoatExcelPro.exe` dùng vĩnh viễn!

### Cách 2: Chạy 1 cú click trên bất kỳ máy nào có sẵn Python
1. Giải nén thư mục.
2. Nhấp đúp chuột vào file **`build_exe.bat`**.
3. Sau 1 phút, lấy file **`dist\InHangLoatExcelPro.exe`** mang sang máy của bạn dùng trực tiếp (không cần Python).

## Yêu cầu trên máy sử dụng `InHangLoatExcelPro.exe`
- Hệ điều hành Windows 10 / 11.
- Đã cài sẵn **Microsoft Excel** (để xử lý chuẩn xác định dạng `.xlsb`, `.xls`, `.xlsm`, `.xlsx`, vùng in `PrintArea` và tỷ lệ thu phóng gốc).
