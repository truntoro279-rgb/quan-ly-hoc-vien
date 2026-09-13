QUAN LY HOC VIEN - BO APP

A. CHẠY GIAO DIỆN
- Upload toàn bộ thư mục này lên Netlify.
- index.html là giao diện chính.
- manifest.json hỗ trợ cài như app.

B. KẾT NỐI GOOGLE SHEETS
1. Mở Google Sheet dữ liệu học viên.
2. Extensions > Apps Script.
3. Dán code trong gas_backend.gs vào Code.gs.
4. Thay DAN_ID_GOOGLE_SHEET_VAO_DAY bằng ID Google Sheet.
5. Deploy > New deployment > Web app.
6. Execute as: Me.
7. Who has access: Anyone.
8. Copy URL /exec.
9. Mở index.html, tìm API_URL='' và dán URL vào giữa dấu nháy.
10. Upload lại app lên Netlify.

C. TÊN SHEET
- Sheet học viên mặc định: ALL.
- Backend sẽ tự tạo VE_QUE và DIEM_DANH nếu chưa có.

D. LƯU Ý
- Cần bảo vệ Google Sheet/backend bằng quyền truy cập phù hợp khi đưa vào sử dụng thật.
- Bản này là nền tảng chức năng; có thể mở rộng đăng nhập, phân quyền, ảnh hồ sơ, thống kê, xuất Excel/PDF.
