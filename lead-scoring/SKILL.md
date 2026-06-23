---
name: ckm:lead-scoring
description: "Chấm điểm và phân loại khách hàng tiềm năng ngành Bất động sản dựa trên mô tả nhu cầu chi tiết, xuất báo cáo Excel với dashboard trực quan và nhật ký chất lượng dữ liệu."
argument-hint: "--url <google_sheet_url_or_local_file> [--output <output_path>]"
license: MIT
metadata:
  author: claudekit
  version: "1.0.0"
---

# Lead Scoring System - Hệ thống Chấm điểm Khách hàng Tiềm năng Bất động sản

Tự động hóa việc phân loại và xếp hạng chất lượng khách hàng tiềm năng dựa trên nội dung mô tả nhu cầu của họ. Hệ thống giúp bộ phận Kinh doanh/Marketing nhanh chóng lọc ra nhóm khách VIP cần chăm sóc ngay lập tức và loại bỏ nhóm khách Rác để tránh lãng phí nguồn lực.

## Khi Nào Kích Hoạt

- Người dùng yêu cầu phân tích, lọc, hoặc chấm điểm danh sách khách hàng từ Google Sheets hoặc tệp Excel/CSV cục bộ.
- Người dùng cung cấp danh sách mô tả nhu cầu bất động sản và cần phân nhóm VIP, Bình thường, Rác.

## Quy Trình Xử Lý

### Bước 1: Thu Thập Dữ Liệu
1. **Lấy từ Google Sheets**: Truy xuất dữ liệu từ đường dẫn Google Sheets công khai (thông qua định dạng xuất CSV tiện lợi) hoặc qua API gspread (nếu có cấu hình quyền).
2. **Kiểm Tra Cột**: Đảm bảo tệp nguồn chứa cột `nhu_cau_mo_ta` (chứa mô tả nhu cầu khách hàng) và `sdt` (số điện thoại) cùng `ten_khach` (tên khách hàng).

### Bước 2: Chấm Điểm Theo Tiêu Chí
Điểm xuất phát ban đầu của mỗi khách hàng là **50 điểm**.

1. **Cộng 50 điểm (Khách hàng VIP)** nếu phát hiện ít nhất một tiêu chí VIP:
   - Ngân sách lớn (>= 20 tỷ hoặc "tài chính mạnh", "không thành vấn đề").
   - Loại hình cao cấp ("Biệt thự đơn lập", "Penthouse", "Shophouse mặt đường lớn", "Quỹ đất công nghiệp", "Sàn văn phòng diện tích lớn").
   - Vị trí đắc địa ("Quận 1", "Ven sông", "Vinhomes Ocean Park", "Phú Mỹ Hưng").
   - Đối tượng VIP ("Chủ doanh nghiệp", "Nhà đầu tư chuyên nghiệp", "Mua sỉ", "Mua số lượng lớn").
   - Tính cấp thiết & Minh bạch ("Pháp lý chuẩn 100%", "Sổ hồng riêng", "gặp trực tiếp chủ đầu tư để đàm phán").

2. **Trừ 50 điểm (Khách hàng Rác)** nếu phát hiện ít nhất một tiêu chí tiêu cực:
   - Yêu cầu phi thực tế (Giá thấp vô lý: VD: mua nhà Quận 1 chỉ 1-2 tỷ, nhà trung tâm có sân vườn hồ bơi chỉ vài trăm triệu).
   - Không có nhu cầu ("Nhầm số", "Không có nhu cầu", "Dữ liệu cũ", "Nhầm ngành").
   - Không thiện chí ("Hỏi giá cho vui", "Chưa có ý định mua", "Thái độ không hợp tác").
   - Spam/Quảng cáo ("Bảo hiểm", "Vay vốn", "Mời chào dịch vụ").
   - Liên lạc lỗi ("Thuê bao", "Gọi nhiều lần không bắt máy", "Không phản hồi Zalo").

3. **Giữ nguyên 50 điểm (Khách hàng Bình thường)** nếu không thuộc hai nhóm trên (VD: mua căn hộ/nhà phố tầm trung 3-10 tỷ, cần vay ngân hàng, hoặc có nhu cầu thực nhưng cần tư vấn thêm).

### Bước 3: Xuất Báo Cáo
Xuất báo cáo kết quả ra tệp Excel định dạng sang trọng chứa:
- `Summary_Dashboard`: Tóm tắt số lượng, tỷ lệ các nhóm khách hàng, điểm trung bình và biểu đồ cơ cấu.
- `Lead_Scores_Details`: Chi tiết điểm số, phân loại (VIP, Bình thường, Rác) kèm lý do giải trình chi tiết cho từng trường hợp.
- `Data_Integrity_Log`: Danh sách các bản ghi bị lỗi dữ liệu (như thiếu số điện thoại, thiếu mô tả nhu cầu) được tự động phát hiện và xử lý.

---

## Hướng Dẫn Chạy Script

Chạy script chấm điểm bằng Python:
```bash
python "C:\Users\HOME\.gemini\config\skills\lead-scoring\scripts\score_leads.py" ^
  --url "https://docs.google.com/spreadsheets/d/1EaHaNMNmqz2Yy-3DpaNktbi4Ii0vRvaz4lwjpL71zYg/edit?usp=sharing" ^
  --output "C:\Users\HOME\Documents\Thực hành bài 7\Lead_Scoring_Report.xlsx"
```

### Tham Số Đầu Vào:
- `--url`: Link Google Sheet hoặc đường dẫn tới file CSV/Excel nội bộ. (Bắt buộc)
- `--output`: Đường dẫn lưu file Excel báo cáo đầu ra. (Mặc định: `Lead_Scoring_Report.xlsx` ở thư mục hiện tại)
