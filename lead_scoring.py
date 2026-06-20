#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import re
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# 1. CONFIGURATION
CREDENTIALS_FILE = "credentials.json"
DEFAULT_SHEET_NAME = "Sheet1"

# Scopes for Google Sheets and Google Drive APIs
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_sheets_client():
    """Initializes and returns the Google Sheets API client using service account credentials."""
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            f"❌ Không tìm thấy tệp '{CREDENTIALS_FILE}'. "
            "Vui lòng đảm bảo tệp này nằm trong cùng thư mục với script."
        )
    
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

# 2. RULE-BASED SCORING ENGINE
def evaluate_lead(demand_text):
    """
    Evaluates a lead based on their demand description.
    Returns: (score, reason_str, classification)
    """
    if not demand_text or not isinstance(demand_text, str):
        return 50, "Không có thông tin nhu cầu", "TRUNG BÌNH"
    
    demand_lower = demand_text.lower()
    score = 50  # Base score
    reasons = []
    
    # --- VIP SIGNALS (ADD 50 POINTS) ---
    vip_signals = []
    
    # 1. Budget >= 20 Billion
    billion_matches = re.findall(r'(\d+)\s*tỷ', demand_lower)
    has_large_budget = False
    for match in billion_matches:
        try:
            if int(match) >= 20:
                has_large_budget = True
                break
        except ValueError:
            continue
            
    if has_large_budget or any(kw in demand_lower for kw in ["tài chính mạnh", "không thành vấn đề", "ngân sách lớn"]):
        vip_signals.append("Ngân sách lớn (>= 20 tỷ)")
        
    # 2. Premium type of property
    vip_types = [
        "biệt thự đơn lập", 
        "penthouse", 
        "shophouse mặt đường lớn", 
        "quỹ đất công nghiệp", 
        "sàn văn phòng diện tích lớn", 
        "sàn văn phòng lớn"
    ]
    for v_type in vip_types:
        if v_type in demand_lower:
            vip_signals.append(f"Loại hình cao cấp ({v_type.title()})")
            break
            
    # 3. Prime location
    vip_locations = ["quận 1", "ven sông", "vinhomes ocean park", "phú mỹ hưng", "thảo điền"]
    for loc in vip_locations:
        if loc in demand_lower:
            vip_signals.append(f"Vị trí đắc địa ({loc.title()})")
            break
            
    # 4. VIP client status
    vip_clients = ["chủ doanh nghiệp", "nhà đầu tư chuyên nghiệp", "mua sỉ", "mua số lượng lớn"]
    for client in vip_clients:
        if client in demand_lower:
            vip_signals.append(f"Khách hàng VIP ({client.title()})")
            break
            
    # 5. Urgency & transparency
    urgency_signals = ["pháp lý chuẩn", "sổ hồng riêng", "gặp trực tiếp chủ đầu tư", "làm việc trực tiếp chủ đầu tư"]
    for sig in urgency_signals:
        if sig in demand_lower:
            vip_signals.append(f"Tính cấp thiết/Minh bạch ({sig.title()})")
            break

    # --- TRASH SIGNALS (SUBTRACT 50 POINTS) ---
    trash_signals = []
    
    # 1. Unrealistic requirements (e.g. Q1 house under 3 billion)
    if "quận 1" in demand_lower or "q1" in demand_lower:
        under_3b_match = re.search(r'(\d+)\s*tỷ', demand_lower)
        if under_3b_match:
            try:
                if int(under_3b_match.group(1)) <= 2:
                    trash_signals.append("Yêu cầu phi thực tế (Nhà Q1 giá dưới 3 tỷ)")
            except ValueError:
                pass
    if "trung tâm" in demand_lower and any(kw in demand_lower for kw in ["vài trăm triệu", "500 triệu", "700 triệu"]):
        trash_signals.append("Yêu cầu phi thực tế (Nhà trung tâm giá quá rẻ)")
        
    # 2. No real demand
    no_need_kws = ["nhầm số", "không có nhu cầu", "dữ liệu cũ", "nhầm ngành", "gọi sai người"]
    for kw in no_need_kws:
        if kw in demand_lower:
            trash_signals.append(f"Không có nhu cầu ({kw})")
            break
            
    # 3. Lack of cooperation
    uncooperative_kws = ["hỏi giá cho vui", "chưa có ý định mua", "thái độ không hợp tác", "cúp máy", "không hợp tác"]
    for kw in uncooperative_kws:
        if kw in demand_lower:
            trash_signals.append("Khách hàng thiếu thiện chí")
            break
            
    # 4. Spam / cross-selling
    spam_kws = ["bảo hiểm", "vay vốn", "mời chào", "gửi tiết kiệm", "quảng cáo"]
    for kw in spam_kws:
        if kw in demand_lower:
            trash_signals.append(f"Spam/Quảng cáo ({kw.title()})")
            break
            
    # 5. Connection issues
    connection_errors = ["thuê bao", "không bắt máy", "gọi nhiều lần không được", "không phản hồi zalo"]
    for err in connection_errors:
        if err in demand_lower:
            trash_signals.append("Lỗi liên lạc")
            break

    # Calculate final score and compile reasons
    if trash_signals:
        score -= 50
        reasons.extend(trash_signals)
    if vip_signals:
        score += 50
        reasons.extend(vip_signals)
        
    # Bound score between 0 and 100
    score = max(0, min(100, score))
    
    # Classify based on score
    if score >= 100:
        classification = "VIP / SIÊU TIỀM NĂNG"
    elif score <= 0:
        classification = "RÁC / KHÔNG TIỀM NĂNG"
    else:
        classification = "TRUNG BÌNH"
        
    reason_str = "; ".join(reasons) if reasons else "Nhu cầu tiêu chuẩn / Tầm trung"
    return score, reason_str, classification

# 3. PROCESSING PIPELINE
def process_lead_scoring(spreadsheet_url_or_id, sheet_name=DEFAULT_SHEET_NAME):
    print("🚀 Đang kết nối với Google Sheet...")
    client = get_sheets_client()
    
    try:
        if "docs.google.com" in spreadsheet_url_or_id:
            sh = client.open_by_url(spreadsheet_url_or_id)
        else:
            sh = client.open_by_key(spreadsheet_url_or_id)
    except Exception as e:
        print(f"❌ Không thể mở Google Sheet. Lỗi: {e}")
        print("👉 Vui lòng đảm bảo rằng bạn đã chia sẻ quyền Editor cho email Service Account:")
        print("   aib5-86@robotic-epoch-499804-d0.iam.gserviceaccount.com")
        sys.exit(1)
        
    try:
        worksheet = sh.worksheet(sheet_name)
    except Exception as e:
        print(f"❌ Không tìm thấy sheet có tên '{sheet_name}'. Lỗi: {e}")
        sys.exit(1)
        
    # Retrieve all records from the worksheet
    data = worksheet.get_all_records()
    if not data:
        print("⚠️ Sheet hiện tại không có dữ liệu hoặc tiêu đề không đúng cấu trúc.")
        sys.exit(1)
        
    df = pd.DataFrame(data)
    print(f"📊 Đã tải thành công {len(df)} dòng dữ liệu.")
    
    # Check for required headers
    required_cols = ["Họ tên", "Số điện thoại", "Mô tả nhu cầu"]
    for col in required_cols:
        if col not in df.columns:
            print(f"❌ Thiếu cột bắt buộc '{col}' trong Google Sheet. Hãy kiểm tra lại dòng tiêu đề đầu tiên.")
            sys.exit(1)
            
    # Initialize output columns if they don't exist
    if "Điểm số" not in df.columns:
        df["Điểm số"] = ""
    if "Phân loại" not in df.columns:
        df["Phân loại"] = ""
    if "Lý do chấm điểm" not in df.columns:
        df["Lý do chấm điểm"] = ""
        
    # Run the lead scoring logic
    print("🧠 Đang phân tích dữ liệu và chấm điểm...")
    vip_count = 0
    trash_count = 0
    
    for idx, row in df.iterrows():
        demand = row["Mô tả nhu cầu"]
        score, reason, classification = evaluate_lead(demand)
        
        df.at[idx, "Điểm số"] = score
        df.at[idx, "Phân loại"] = classification
        df.at[idx, "Lý do chấm điểm"] = reason
        
        if classification == "VIP / SIÊU TIỀM NĂNG":
            vip_count += 1
        elif classification == "RÁC / KHÔNG TIỀM NĂNG":
            trash_count += 1

    # Save results back to Google Sheets
    print("💾 Đang lưu kết quả lại Google Sheets...")
    headers = worksheet.row_values(1)
    
    def get_or_create_col_idx(col_name):
        if col_name in headers:
            return headers.index(col_name) + 1
        else:
            new_col_idx = len(headers) + 1
            worksheet.update_cell(1, new_col_idx, col_name)
            headers.append(col_name)
            return new_col_idx

    score_col_idx = get_or_create_col_idx("Điểm số")
    class_col_idx = get_or_create_col_idx("Phân loại")
    reason_col_idx = get_or_create_col_idx("Lý do chấm điểm")
    
    # Batch update to prevent rate limits
    score_cells = []
    class_cells = []
    reason_cells = []
    
    for idx, row in df.iterrows():
        row_num = idx + 2  # 1-based and skips headers
        score_cells.append(gspread.Cell(row=row_num, col=score_col_idx, value=row["Điểm số"]))
        class_cells.append(gspread.Cell(row=row_num, col=class_col_idx, value=row["Phân loại"]))
        reason_cells.append(gspread.Cell(row=row_num, col=reason_col_idx, value=row["Lý do chấm điểm"]))
        
    worksheet.update_cells(score_cells)
    worksheet.update_cells(class_cells)
    worksheet.update_cells(reason_cells)
    
    print("\n" + "="*40)
    print("🎉 Hoàn tất chấm điểm khách hàng tiềm năng!")
    print(f"📊 Tổng số leads đã xử lý: {len(df)}")
    print(f"⭐ VIP / Siêu tiềm năng: {vip_count}")
    print(f"🗑️ Rác / Không tiềm năng: {trash_count}")
    print(f"📈 Trung bình: {len(df) - vip_count - trash_count}")
    print("="*40)

if __name__ == "__main__":
    # Check if arguments are provided via CLI
    if len(sys.argv) > 1:
        url_or_id = sys.argv[1]
        sheet = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_SHEET_NAME
        process_lead_scoring(url_or_id, sheet)
    else:
        print("💡 Hướng dẫn chạy:")
        print("   python lead_scoring.py <SPREADSHEET_URL_OR_ID> [SHEET_NAME]")
        print("\nHoặc nhập trực tiếp thông tin bên dưới:")
        
        try:
            url_or_id = input("Nhập URL hoặc ID Google Sheet: ").strip()
            if not url_or_id:
                print("❌ URL hoặc ID không được để trống.")
                sys.exit(1)
            sheet = input(f"Nhập tên sheet (Bấm Enter để mặc định '{DEFAULT_SHEET_NAME}'): ").strip()
            if not sheet:
                sheet = DEFAULT_SHEET_NAME
            process_lead_scoring(url_or_id, sheet)
        except KeyboardInterrupt:
            print("\n👋 Đã hủy chương trình.")
            sys.exit(0)
