# Prompt gửi cho AI khác khi đính kèm 4 file NO/MO

Bạn là chuyên gia phân tích dữ liệu SAP PM, dashboard vận hành bảo trì và Python/Streamlit. Tôi sẽ đính kèm 4 file Excel xuất từ SAP nhưng định dạng thực tế có thể là MHT/HTML table:

- NO.MHT: dữ liệu ZTC NO, kỳ cũ, từ 01-01-2026 đến 19-04-2026
- MO.MHT: dữ liệu ZTC MO, kỳ cũ, từ 01-01-2026 đến 19-04-2026
- NO 2.MHT: dữ liệu ZTC NO, kỳ mới, từ 01-01-2026 đến 26-04-2026
- MO 2.MHT: dữ liệu ZTC MO, kỳ mới, từ 01-01-2026 đến 26-04-2026

Yêu cầu của bạn là đọc hiểu dữ liệu, không được đoán khi chưa kiểm tra bằng dữ liệu. Hãy làm đầy đủ các bước sau:

## 1. Đọc và chuẩn hóa file

1. Đọc 4 file dưới dạng bảng dữ liệu.
2. Nếu file là MHT/HTML, hãy dùng parser HTML table, không giả định đây là xlsx chuẩn.
3. Xác định dòng header thật.
4. Với file MO, loại các dòng tổng/nhóm không phải lệnh thực. Lệnh thực là dòng có `Lệnh sửa chữa` hợp lệ.
5. Chuẩn hóa lỗi mã hóa tiếng Việt nếu có, ví dụ `Ðã đóng` thành `Đã đóng`.
6. Chuẩn hóa mã ID như `Mã sự cố`, `Lệnh sửa chữa`, `Thiết bị`, `Mã kế hoạch` thành chuỗi, không để mất số hoặc biến thành `.0`.
7. Chuẩn hóa ngày tháng thành date.

## 2. Hiểu ý nghĩa nghiệp vụ

Với file NO:
- 1 dòng tương ứng 1 sự cố/thông báo.
- Khóa chính: `Mã sự cố`.
- Cột liên kết sang MO: `Lệnh sửa chữa`.
- Cột trạng thái: `Tên T.thái`.
- Cột cửa hàng/khu vực ưu tiên: `Tên KV chức năng`, sau đó `Tên PB S/dụng`, sau đó `Tên đơn vị`.
- Cột nhóm công việc/thiết bị: `Tên nhóm T/bị`, `Tên tổ đội S/ chữa`, `Tên BPhận lập KH`.

Với file MO:
- 1 dòng thực tương ứng 1 lệnh sửa chữa/bảo trì.
- Khóa chính: `Lệnh sửa chữa`.
- Cột liên kết về NO: `Mã sự cố`.
- Cột phân loại nguồn: `Lệnh S/ chữa từ`, thường gồm `Từ sự cố` hoặc `Từ kế hoạch`.
- MO từ kế hoạch thường có `Mã kế hoạch` và không có `Mã sự cố`.
- MO từ sự cố thường có `Mã sự cố` và có thể liên kết với NO.

## 3. Kiểm tra cấu trúc từng file

Hãy báo cáo:
- Số bảng tìm thấy trong mỗi file.
- Số dòng tổng và số dòng dữ liệu thực.
- Số cột.
- Danh sách cột quan trọng.
- Số dòng bị loại vì là dòng tổng/nhóm.
- Có mã khóa trùng hay không.
- Có thiếu cột bắt buộc hay không.

## 4. Tổng quan dữ liệu hiện tại

Với từng file NO:
- Tổng số sự cố.
- Số sự cố theo trạng thái: `Đã đóng`, `Khởi tạo`, `Phê duyệt`, `Thực hiện`, `Hoàn thành`, `Từ chối`, trạng thái khác nếu có.
- Số NO có `Lệnh sửa chữa`.
- Số NO chưa có `Lệnh sửa chữa`.
- Top cửa hàng có nhiều NO mở.
- Top nhóm công việc/nhóm thiết bị có nhiều NO mở.
- Aging NO mở: 0-7 ngày, 8-14 ngày, 15-30 ngày, trên 30 ngày.

Với từng file MO:
- Tổng số lệnh thực.
- Số MO theo trạng thái.
- Số MO theo nguồn: `Từ kế hoạch`, `Từ sự cố`.
- Số MO có `Mã sự cố`.
- Số MO có `Mã kế hoạch`.

## 5. Liên kết NO-MO

Liên kết dữ liệu bằng:

- `NO.Mã sự cố` ↔ `MO.Mã sự cố`
- kiểm tra chéo thêm `NO.Lệnh sửa chữa` ↔ `MO.Lệnh sửa chữa`

Hãy phân loại:
- `NO có MO khớp`: khớp cả `Mã sự cố` và `Lệnh sửa chữa`.
- `NO có lệnh nhưng thiếu trong MO`: NO có `Lệnh sửa chữa` nhưng không thấy trong MO.
- `NO chưa có lệnh sửa chữa`: NO chưa có `Lệnh sửa chữa`.
- `MO từ kế hoạch`: MO có `Mã kế hoạch`, không liên quan NO.
- `MO từ sự cố nhưng không thấy NO`: MO có `Mã sự cố` nhưng không tìm được NO.

Hãy liệt kê số lượng từng nhóm và nêu các mã bất thường quan trọng.

## 6. So sánh kỳ cũ và kỳ mới

So sánh NO.MHT với NO 2.MHT bằng `Mã sự cố`:
- Sự cố mới phát sinh.
- Sự cố biến mất khỏi file mới nếu có.
- Sự cố chuyển trạng thái.
- Sự cố đứng yên trạng thái.
- Sự cố đã đóng trong kỳ.
- Delta số lượng theo từng trạng thái.
- Delta NO mở theo cửa hàng.
- Delta NO mở theo nhóm công việc.

So sánh MO.MHT với MO 2.MHT bằng `Lệnh sửa chữa`:
- MO mới phát sinh.
- MO biến mất khỏi file mới nếu có.
- MO chuyển trạng thái nếu có.
- Kiểm tra hai file MO có giống hệt nhau không bằng hash hoặc so sánh dữ liệu.

## 7. Sinh insight điều tra

Tạo danh sách các dòng cần điều tra, gồm:
- NO còn mở nhưng chưa sinh MO.
- NO có lệnh sửa chữa nhưng không tìm thấy trong file MO.
- NO đã có MO nhưng NO vẫn còn mở.
- MO hoàn thành kỹ thuật nhưng NO chưa đóng.
- NO hoàn thành nhưng chưa đóng.
- NO tồn trên 14 ngày.
- NO tồn trên 30 ngày.
- NO ở trạng thái khởi tạo quá 7 ngày.
- Cửa hàng có NO mở tăng mạnh so với tuần trước.
- Nhóm công việc có NO mở tăng mạnh so với tuần trước.
- Thiết bị/cửa hàng có lỗi lặp lại.

Với mỗi dòng điều tra, hãy có cột `Cờ điều tra` giải thích lý do.

## 8. Kết quả cuối cùng cần tạo

Hãy trả lời bằng tiếng Việt và tạo kết quả cuối cùng gồm:

1. Bản tóm tắt executive summary cho sếp.
2. Bảng KPI tổng quan tuần hiện tại.
3. Bảng delta so với tuần trước.
4. Bảng top cửa hàng/nhóm công việc biến động mạnh.
5. Bảng danh sách NO/MO cần điều tra.
6. Nhận xét về chất lượng dữ liệu, ví dụ file MO tuần mới có trùng file MO tuần cũ hay không.
7. Nếu được yêu cầu viết code, hãy tạo Streamlit app có:
   - Upload NO/MO tuần hiện tại.
   - Upload NO/MO tuần trước để so sánh.
   - Tab Tổng quan.
   - Tab So sánh tuần.
   - Tab Điều tra chi tiết.
   - Tab Dữ liệu gốc/chuẩn hóa.
   - Nút download Excel báo cáo.
   - Cảnh báo khi file MO tuần này giống tuần trước hoặc thiếu cột bắt buộc.

Không chỉ mô tả chung chung. Hãy tính toán bằng dữ liệu thật từ file, đưa ra con số cụ thể và nêu rõ giả định nếu có.
