---
name: ckm:lead-scoring-real-estate
description: "Hướng dẫn AI kết nối Google Sheets, đọc dữ liệu khách hàng (leads), tự động chấm điểm và phân loại khách hàng tiềm năng ngành Bất động sản theo bộ tiêu chí chuẩn."
argument-hint: "<spreadsheet_url_or_id> [sheet_name]"
license: MIT
metadata:
  author: Antigravity
  version: "1.0.0"
---

# Hướng Dẫn Chấm Điểm Khách Hàng Tiềm Năng (Lead Scoring) - Ngành Bất Động Sản

Tài liệu này hướng dẫn cách xây dựng và vận hành một hệ thống chấm điểm tự động cho các khách hàng tiềm năng (leads) trong lĩnh vực Bất động sản. Hệ thống sẽ kết nối với **Google Sheets** bằng tài khoản dịch vụ (Service Account), thu thập thông tin khách hàng, phân tích nhu cầu bằng AI/Regex để chấm điểm, và ghi kết quả ngược lại Google Sheets.

---

## 📌 Chuẩn Bị & Cấu Hình Kết Nối

Để hệ thống hoạt động, bạn cần cấu hình quyền truy cập Google Sheets thông qua Service Account có sẵn trong tệp [credentials.json](file:///c:/Users/HOME/Desktop/Agentic%20AI/Buoi%207/credentials.json).

### 1. Chia sẻ quyền truy cập Google Sheet
1. Mở file Google Sheet chứa dữ liệu khách hàng của bạn.
2. Nhấp vào nút **Share** (Chia sẻ) ở góc trên bên phải.
3. Thêm email của Service Account dưới đây với quyền **Editor** (Người chỉnh sửa):
   ```text
   aib5-86@robotic-epoch-499804-d0.iam.gserviceaccount.com
   ```
4. Nhấn **Send** (Gửi) để hoàn tất.

### 2. Cài đặt thư viện Python cần thiết
Chạy lệnh sau để cài đặt các thư viện kết nối Google Sheets và xử lý dữ liệu:
```bash
pip install gspread google-auth pandas openpyxl
```

---

## 📋 Cấu Trúc Bảng Dữ Liệu Khách Hàng (Google Sheet)

Bảng tính Google Sheet của bạn cần có các cột tối thiểu sau:
- **Họ tên**: Tên của khách hàng.
- **Số điện thoại**: Số điện thoại liên hệ.
- **Mô tả nhu cầu**: Nội dung ghi chú hoặc mô tả nhu cầu chi tiết của khách hàng (đây là cột chính để AI phân tích).
- **Điểm số**: Cột để hệ thống ghi điểm (Score).
- **Phân loại**: Cột để hệ thống phân loại khách hàng (VIP, Trung bình, Rác).
- **Lý do chấm điểm**: Cột ghi nhận nguyên nhân cộng/trừ điểm để nhân viên tư vấn dễ theo dõi.

---

## ⚖️ Quy Tắc Chấm Điểm (Theo Tiêu Chí Chuẩn)

Dựa trên tài liệu [tieu_chi_cham_diem.txt](file:///c:/Users/HOME/Desktop/Agentic%20AI/Buoi%207/tieu_chi_cham_diem.txt), điểm mặc định ban đầu của mỗi Lead là **0 điểm** (hoặc **50 điểm** tùy cấu hình cơ sở, ở đây sử dụng thang điểm mặc định bắt đầu từ `50` và cộng/trừ để ra điểm cuối cùng).

### 🟢 1. Cộng 50 điểm (Khách hàng VIP / Siêu tiềm năng)
Cần nhận diện các từ khóa hoặc ngữ cảnh sau trong cột **Mô tả nhu cầu**:
- **Ngân sách lớn**: Đề cập số tiền cụ thể **từ 20 tỷ trở lên** hoặc chứa từ khóa: `tài chính mạnh`, `không thành vấn đề`, `ngân sách lớn`, `trên 20 tỷ`, `20-30 tỷ`.
- **Loại hình cao cấp**: Tìm kiếm các từ khóa: `Biệt thự đơn lập`, `Penthouse`, `Shophouse mặt đường lớn`, `Quỹ đất công nghiệp`, `Sàn văn phòng diện tích lớn`.
- **Vị trí đắc địa**: Yêu cầu các khu vực: `Quận 1`, `Ven sông`, `Vinhomes Ocean Park`, `Phú Mỹ Hưng`, `Thảo Điền`.
- **Đối tượng khách hàng**: Đề cập hoặc có dấu hiệu là: `Chủ doanh nghiệp`, `Nhà đầu tư chuyên nghiệp`, `Mua sỉ`, `Mua số lượng lớn`.
- **Tính cấp thiết & Minh bạch**: Đòi hỏi pháp lý cao hoặc muốn làm việc trực tiếp: `Pháp lý chuẩn 100%`, `Sổ hồng riêng`, `gặp trực tiếp chủ đầu tư`, `mua ngay trong tháng`.

### 🔴 2. Trừ 50 điểm (Khách hàng Rác / Không tiềm năng)
Cần phát hiện các dấu hiệu sau để trừ điểm hoặc loại bỏ trực tiếp:
- **Yêu cầu phi thực tế**: Mua bất động sản giá rẻ vô lý so với thị trường (VD: `Nhà Quận 1 giá 1-2 tỷ`, `nhà trung tâm có sân vườn hồ bơi giá vài trăm triệu`, `mua đất nền trung tâm 500 triệu`).
- **Không có nhu cầu thực tế**: Từ khóa: `Nhầm số`, `Không có nhu cầu`, `Dữ liệu cũ`, `Nhầm ngành`, `Gọi sai người`.
- **Thiếu thiện chí**: Từ khóa: `Hỏi giá cho vui`, `Chưa có ý định mua`, `Thái độ không hợp tác`, `Cúp máy`, `Không hợp tác`.
- **Spam / Quảng cáo chéo**: Chứa nội dung quảng cáo dịch vụ khác như: `Bảo hiểm`, `Vay vốn`, `Mới chào dịch vụ`, `Gửi tiết kiệm`.
- **Thông tin liên lạc lỗi**: Trạng thái: `Thuê bao`, `Không bắt máy`, `Gọi nhiều lần không được`, `Không phản hồi zalo`.

### 🟡 3. Các trường hợp trung bình (Giữ nguyên điểm / Cộng ít)
- Khách hàng tìm mua chung cư, nhà phố tầm trung (giá trị từ `3 - 10 tỷ`).
- Khách hàng cần hỗ trợ vay ngân hàng, đang cân nhắc chính sách tài chính.
- Khách hàng có nhu cầu thực nhưng cần tư vấn sâu hơn về pháp lý hoặc vị trí.

---

## 🛠️ Code Python Mẫu: Chấm Điểm Tự Động (Hybrid: Rule-based & LLM)

Dưới đây là mã nguồn Python đầy đủ giúp tự động hóa toàn bộ quy trình từ kết nối Google Sheets, chấm điểm và cập nhật dữ liệu.

```python
import os
import re
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# 1. CẤU HÌNH KẾT NỐI GOOGLE SHEET
CREDENTIALS_FILE = "credentials.json" # Đảm bảo file này nằm cùng thư mục chạy script
SPREADSHEET_KEY_OR_URL = "ĐIỀN_URL_HOẶC_ID_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY"
SHEET_NAME = "Sheet1" # Tên sheet cần chấm điểm

# Khởi tạo kết nối Google Sheets API
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_sheets_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    return client

# 2. HÀM CHẤM ĐIỂM DỰA TRÊN LUẬT (RULE-BASED ENGINE)
def evaluate_lead(demand_text):
    if not demand_text or not isinstance(demand_text, str):
        return 50, "Không có thông tin nhu cầu", "Trung bình"
    
    demand_lower = demand_text.lower()
    score = 50 # Điểm gốc ban đầu
    reasons = []
    
    # --- CÁC TIÊU CHÍ CỘNG ĐIỂM VIP (+50) ---
    vip_signals = []
    
    # Kiểm tra ngân sách lớn >= 20 tỷ
    billion_matches = re.findall(r'(\d+)\s*tỷ', demand_lower)
    has_large_budget = False
    for match in billion_matches:
        if int(match) >= 20:
            has_large_budget = True
            break
    if has_large_budget or any(kw in demand_lower for kw in ["tài chính mạnh", "không thành vấn đề", "ngân sách lớn"]):
        vip_signals.append("Ngân sách lớn (>= 20 tỷ)")
        
    # Loại hình cao cấp
    vip_types = ["biệt thự đơn lập", "penthouse", "shophouse mặt đường lớn", "quỹ đất công nghiệp", "sàn văn phòng diện tích lớn", "sàn văn phòng lớn"]
    for v_type in vip_types:
        if v_type in demand_lower:
            vip_signals.append(f"Loại hình cao cấp ({v_type.title()})")
            break
            
    # Vị trí đắc địa
    vip_locations = ["quận 1", "ven sông", "vinhomes ocean park", "phú mỹ hưng", "thảo điền"]
    for loc in vip_locations:
        if loc in demand_lower:
            vip_signals.append(f"Vị trí đắc địa ({loc.title()})")
            break
            
    # Đối tượng VIP
    vip_clients = ["chủ doanh nghiệp", "nhà đầu tư chuyên nghiệp", "mua sỉ", "mua số lượng lớn"]
    for client in vip_clients:
        if client in demand_lower:
            vip_signals.append(f"Khách hàng VIP ({client.title()})")
            break
            
    # Tính cấp thiết & Minh bạch
    urgency_signals = ["pháp lý chuẩn", "sổ hồng riêng", "gặp trực tiếp chủ đầu tư", "làm việc trực tiếp chủ đầu tư"]
    for sig in urgency_signals:
        if sig in demand_lower:
            vip_signals.append(f"Tính cấp thiết/Minh bạch ({sig.title()})")
            break

    # --- CÁC TIÊU CHÍ TRỪ ĐIỂM RÁC (-50) ---
    trash_signals = []
    
    # Yêu cầu phi thực tế (Ví dụ: nhà Q1 giá dưới 3 tỷ)
    if "quận 1" in demand_lower or "q1" in demand_lower:
        under_3b_match = re.search(r'(\d+)\s*tỷ', demand_lower)
        if under_3b_match and int(under_3b_match.group(1)) <= 2:
            trash_signals.append("Yêu cầu phi thực tế (Nhà Q1 giá dưới 3 tỷ)")
    if "trung tâm" in demand_lower and any(kw in demand_lower for kw in ["vài trăm triệu", "500 triệu", "700 triệu"]):
        trash_signals.append("Yêu cầu phi thực tế (Nhà trung tâm giá quá rẻ)")
        
    # Không có nhu cầu
    no_need_kws = ["nhầm số", "không có nhu cầu", "dữ liệu cũ", "nhầm ngành", "gọi sai người"]
    for kw in no_need_kws:
        if kw in demand_lower:
            trash_signals.append(f"Không có nhu cầu ({kw})")
            break
            
    # Không thiện chí
    uncooperative_kws = ["hỏi giá cho vui", "chưa có ý định mua", "thái độ không hợp tác", "cúp máy", "không hợp tác"]
    for kw in uncooperative_kws:
        if kw in demand_lower:
            trash_signals.append("Khách hàng thiếu thiện chí")
            break
            
    # Spam/Quảng cáo
    spam_kws = ["bảo hiểm", "vay vốn", "mời chào", "gửi tiết kiệm", "quảng cáo"]
    for kw in spam_kws:
        if kw in demand_lower:
            trash_signals.append(f"Spam/Quảng cáo ({kw.title()})")
            break
            
    # Liên lạc lỗi
    connection_errors = ["thuê bao", "không bắt máy", "gọi nhiều lần không được", "không phản hồi zalo"]
    for err in connection_errors:
        if err in demand_lower:
            trash_signals.append("Lỗi liên lạc")
            break

    # Tính toán kết quả cuối cùng
    if trash_signals:
        score -= 50
        reasons.extend(trash_signals)
    if vip_signals:
        score += 50
        reasons.extend(vip_signals)
        
    # Giới hạn thang điểm từ 0 đến 100
    score = max(0, min(100, score))
    
    # Phân loại khách hàng
    if score >= 100:
        classification = "VIP / SIÊU TIỀM NĂNG"
    elif score <= 0:
        classification = "RÁC / KHÔNG TIỀM NĂNG"
    else:
        classification = "TRUNG BÌNH"
        
    reason_str = "; ".join(reasons) if reasons else "Nhu cầu tiêu chuẩn / Tầm trung"
    return score, reason_str, classification

# 3. QUY TRÌNH ĐỌC VÀ GHI ĐÈ GOOGLE SHEET
def process_lead_scoring():
    print("🚀 Đang kết nối với Google Sheet...")
    client = get_sheets_client()
    
    try:
        sh = client.open_by_url(SPREADSHEET_KEY_OR_URL)
    except Exception:
        # Nếu không mở được bằng URL, thử mở bằng ID
        sh = client.open_by_key(SPREADSHEET_KEY_OR_URL)
        
    worksheet = sh.worksheet(SHEET_NAME)
    
    # Lấy toàn bộ dữ liệu dưới dạng danh sách các dict
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    
    print(f"📊 Đã tải thành công {len(df)} dòng dữ liệu.")
    
    # Kiểm tra sự tồn tại của các cột cần thiết
    required_cols = ["Họ tên", "Số điện thoại", "Mô tả nhu cầu"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Thiếu cột bắt buộc '{col}' trong Google Sheet. Hãy kiểm tra lại tiêu đề.")
            
    # Tạo các cột kết quả nếu chưa có
    if "Điểm số" not in df.columns:
        df["Điểm số"] = ""
    if "Phân loại" not in df.columns:
        df["Phân loại"] = ""
    if "Lý do chấm điểm" not in df.columns:
        df["Lý do chấm điểm"] = ""
        
    # Duyệt qua từng dòng và thực hiện chấm điểm
    print("🧠 Đang phân tích dữ liệu và chấm điểm...")
    for idx, row in df.iterrows():
        demand = row["Mô tả nhu cầu"]
        score, reason, classification = evaluate_lead(demand)
        
        df.at[idx, "Điểm số"] = score
        df.at[idx, "Phân loại"] = classification
        df.at[idx, "Lý do chấm điểm"] = reason

    # Ghi dữ liệu đã cập nhật ngược lại Google Sheets
    print("💾 Đang lưu kết quả lại Google Sheets...")
    
    # Xác định vị trí của cột ghi kết quả để tối ưu ghi đè
    headers = worksheet.row_values(1)
    
    # Helper function để lấy chỉ số cột (1-based index) hoặc tạo mới cột
    def get_or_create_col_idx(col_name):
        if col_name in headers:
            return headers.index(col_name) + 1
        else:
            # Append thêm cột mới ở cuối bảng tính
            new_col_idx = len(headers) + 1
            worksheet.update_cell(1, new_col_idx, col_name)
            headers.append(col_name)
            return new_col_idx

    score_col_idx = get_or_create_col_idx("Điểm số")
    class_col_idx = get_or_create_col_idx("Phân loại")
    reason_col_idx = get_or_create_col_idx("Lý do chấm điểm")
    
    # Chuẩn bị dữ liệu ghi theo khối (Batch Update) để tránh vượt quá giới hạn API của Google
    score_cells = []
    class_cells = []
    reason_cells = []
    
    for idx, row in df.iterrows():
        row_num = idx + 2 # Dòng trong Google Sheet là 1-based và bỏ qua tiêu đề (dòng 1)
        score_cells.append(gspread.Cell(row=row_num, col=score_col_idx, value=row["Điểm số"]))
        class_cells.append(gspread.Cell(row=row_num, col=class_col_idx, value=row["Phân loại"]))
        reason_cells.append(gspread.Cell(row=row_num, col=reason_col_idx, value=row["Lý do chấm điểm"]))
        
    # Thực hiện cập nhật hàng loạt
    worksheet.update_cells(score_cells)
    worksheet.update_cells(class_cells)
    worksheet.update_cells(reason_cells)
    
    print("🎉 Hoàn tất! Dữ liệu đã được cập nhật thành công trên Google Sheets.")

if __name__ == "__main__":
    # Điền URL hoặc ID Google Sheet thực tế của bạn trước khi chạy
    if SPREADSHEET_KEY_OR_URL == "ĐIỀN_URL_HOẶC_ID_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY":
        print("⚠️ Vui lòng cập nhật biến SPREADSHEET_KEY_OR_URL bằng URL Google Sheet thực tế.")
    else:
        process_lead_scoring()
```

---

## 🚀 Hướng Dẫn Vận Hành Dành Cho AI

Khi người dùng yêu cầu thực hiện chấm điểm khách hàng tiềm năng:

1. **Bước 1**: Yêu cầu người dùng cung cấp **URL Google Sheet** (hoặc ID) và **Tên Sheet** cần thực hiện chấm điểm.
2. **Bước 2**: Nhắc nhở người dùng thực hiện chia sẻ quyền truy cập Google Sheet cho Service Account: `aib5-86@robotic-epoch-499804-d0.iam.gserviceaccount.com`.
3. **Bước 3**: Tạo một script Python (ví dụ: `run_lead_scoring.py`) trong thư mục làm việc hiện tại, chép toàn bộ code mẫu ở trên vào và thay thế giá trị `SPREADSHEET_KEY_OR_URL` bằng URL của khách hàng cung cấp.
4. **Bước 4**: Thực thi script bằng lệnh:
   ```bash
   python run_lead_scoring.py
   ```
5. **Bước 5**: Kiểm tra kết quả trên Google Sheet và thông báo cho người dùng kèm theo số lượng leads đã chấm điểm thành công, số lượng lead VIP và số lượng lead Rác được phát hiện.
