import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import io
import gspread
from google.oauth2.service_account import Credentials
import altair as alt

# --- PAGE SETUP ---
st.set_page_config(
    page_title="Hệ thống Chấm điểm & Duyệt Khách hàng BĐS",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium Design Look (dark side, elegant spacing, glassmorphism)
st.markdown("""
<style>
    .reportview-container {
        background: #f8f9fa;
    }
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    h1, h2, h3 {
        font-family: 'Outfit', 'Inter', sans-serif;
        color: #1E293B;
    }
    .metric-card {
        background-color: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
        border: 1px solid #E2E8F0;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s ease-in-out;
    }
    .stButton>button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
</style>
""", unsafe_allow_html=True)

# --- CONFIG & CONSTANTS ---
CREDENTIALS_FILE = "credentials.json"
SERVICE_ACCOUNT_EMAIL = "aib5-86@robotic-epoch-499804-d0.iam.gserviceaccount.com"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# --- SESSION STATE INITIALIZATION ---
if "df" not in st.session_state:
    st.session_state.df = None
if "source_type" not in st.session_state:
    st.session_state.source_type = None
if "spreadsheet_url" not in st.session_state:
    st.session_state.spreadsheet_url = ""
if "sheet_name" not in st.session_state:
    st.session_state.sheet_name = "Sheet1"

# --- SCORING ENGINE LOGIC ---
def evaluate_lead(demand_text):
    if not demand_text or not isinstance(demand_text, str):
        return 50, "Không có thông tin nhu cầu", "TRUNG BÌNH"
    
    demand_lower = demand_text.lower()
    score = 50  # Base score
    reasons = []
    
    # --- VIP SIGNALS (+50) ---
    vip_signals = []
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
        
    vip_types = ["biệt thự đơn lập", "penthouse", "shophouse mặt đường lớn", "quỹ đất công nghiệp", "sàn văn phòng diện tích lớn", "sàn văn phòng lớn"]
    for v_type in vip_types:
        if v_type in demand_lower:
            vip_signals.append(f"Loại hình cao cấp ({v_type.title()})")
            break
            
    vip_locations = ["quận 1", "ven sông", "vinhomes ocean park", "phú mỹ hưng", "thảo điền"]
    for loc in vip_locations:
        if loc in demand_lower:
            vip_signals.append(f"Vị trí đắc địa ({loc.title()})")
            break
            
    vip_clients = ["chủ doanh nghiệp", "nhà đầu tư chuyên nghiệp", "mua sỉ", "mua số lượng lớn"]
    for client in vip_clients:
        if client in demand_lower:
            vip_signals.append(f"Khách hàng VIP ({client.title()})")
            break
            
    urgency_signals = ["pháp lý chuẩn", "sổ hồng riêng", "gặp trực tiếp chủ đầu tư", "làm việc trực tiếp chủ đầu tư"]
    for sig in urgency_signals:
        if sig in demand_lower:
            vip_signals.append(f"Tính cấp thiết/Minh bạch ({sig.title()})")
            break

    # --- TRASH SIGNALS (-50) ---
    trash_signals = []
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
        
    no_need_kws = ["nhầm số", "không có nhu cầu", "dữ liệu cũ", "nhầm ngành", "gọi sai người"]
    for kw in no_need_kws:
        if kw in demand_lower:
            trash_signals.append(f"Không có nhu cầu ({kw})")
            break
            
    uncooperative_kws = ["hỏi giá cho vui", "chưa có ý định mua", "thái độ không hợp tác", "cúp máy", "không hợp tác"]
    for kw in uncooperative_kws:
        if kw in demand_lower:
            trash_signals.append("Khách hàng thiếu thiện chí")
            break
            
    spam_kws = ["bảo hiểm", "vay vốn", "mời chào", "gửi tiết kiệm", "quảng cáo"]
    for kw in spam_kws:
        if kw in demand_lower:
            trash_signals.append(f"Spam/Quảng cáo ({kw.title()})")
            break
            
    connection_errors = ["thuê bao", "không bắt máy", "gọi nhiều lần không được", "không phản hồi zalo"]
    for err in connection_errors:
        if err in demand_lower:
            trash_signals.append("Lỗi liên lạc")
            break

    # Final calculations
    if trash_signals:
        score -= 50
        reasons.extend(trash_signals)
    if vip_signals:
        score += 50
        reasons.extend(vip_signals)
        
    score = max(0, min(100, score))
    
    if score >= 100:
        classification = "VIP / SIÊU TIỀM NĂNG"
    elif score <= 0:
        classification = "RÁC / KHÔNG TIỀM NĂNG"
    else:
        classification = "TRUNG BÌNH"
        
    reason_str = "; ".join(reasons) if reasons else "Nhu cầu tiêu chuẩn / Tầm trung"
    return score, reason_str, classification

# --- MOCK DATA ---
def load_mock_data():
    mock_leads = [
        {
            "Họ tên": "Nguyễn Văn Hùng",
            "Số điện thoại": "0912345678",
            "Mô tả nhu cầu": "Tìm kiếm mua biệt thự đơn lập ở Vinhomes Ocean Park, yêu cầu diện tích trên 250m2, ngân sách tầm 35 tỷ đổ lại. Sổ hồng riêng pháp lý chuẩn. Cần gặp trực tiếp chủ nhà đàm phán mua nhanh trong tháng.",
            "Điểm số": "",
            "Phân loại": "",
            "Lý do chấm điểm": "",
            "Trạng thái duyệt": "Chờ duyệt"
        },
        {
            "Họ tên": "Trần Thị Lan",
            "Số điện thoại": "0987654321",
            "Mô tả nhu cầu": "Gọi điện báo nhầm số, tôi không quan tâm đất đai gì hết, dữ liệu cũ rồi đừng gọi lại phiền quá.",
            "Điểm số": "",
            "Phân loại": "",
            "Lý do chấm điểm": "",
            "Trạng thái duyệt": "Chờ duyệt"
        },
        {
            "Họ tên": "Lê Hoàng Nam",
            "Số điện thoại": "0909998887",
            "Mô tả nhu cầu": "Muốn mua nhà mặt tiền ngay trung tâm Quận 1 kinh doanh, diện tích sử dụng rộng rãi tầm 3 tầng trở lên. Mà tài chính chỉ có khoảng 1.5 tỷ thôi. Có nhà nào giới thiệu nhé.",
            "Điểm số": "",
            "Phân loại": "",
            "Lý do chấm điểm": "",
            "Trạng thái duyệt": "Chờ duyệt"
        },
        {
            "Họ tên": "Phạm Minh Đức",
            "Số điện thoại": "0934567890",
            "Mô tả nhu cầu": "Cần tìm mua căn hộ 2 phòng ngủ cho gia đình trẻ tại khu vực quận Bình Thạnh, ngân sách tầm 3.5 tỷ, có hỗ trợ vay ngân hàng 70%. Dự kiến nhận nhà cuối năm nay.",
            "Điểm số": "",
            "Phân loại": "",
            "Lý do chấm điểm": "",
            "Trạng thái duyệt": "Chờ duyệt"
        },
        {
            "Họ tên": "Công ty TNHH Bảo Việt",
            "Số điện thoại": "0243555555",
            "Mô tả nhu cầu": "Chào bạn, bên mình chuyên cung cấp dịch vụ bảo hiểm liên kết ngân hàng và hỗ trợ tài chính doanh nghiệp. Rất mong được hợp tác...",
            "Điểm số": "",
            "Phân loại": "",
            "Lý do chấm điểm": "",
            "Trạng thái duyệt": "Chờ duyệt"
        }
    ]
    return pd.DataFrame(mock_leads)

# --- GOOGLE SHEETS FUNCTIONS ---
def get_sheets_client():
    if not os.path.exists(CREDENTIALS_FILE):
        return None
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception:
        return None

def fetch_from_gsheets(url_or_id, sheet_name):
    client = get_sheets_client()
    if not client:
        raise ConnectionError("Không thể xác thực bằng file credentials.json.")
    
    if "docs.google.com" in url_or_id:
        sh = client.open_by_url(url_or_id)
    else:
        sh = client.open_by_key(url_or_id)
        
    worksheet = sh.worksheet(sheet_name)
    data = worksheet.get_all_records()
    
    if not data:
        return pd.DataFrame(columns=["Họ tên", "Số điện thoại", "Mô tả nhu cầu", "Điểm số", "Phân loại", "Lý do chấm điểm", "Trạng thái duyệt"])
    
    df = pd.DataFrame(data)
    
    # Ensure standard columns are present
    for col in ["Điểm số", "Phân loại", "Lý do chấm điểm", "Trạng thái duyệt"]:
        if col not in df.columns:
            df[col] = ""
            
    # Set default values for Trạng thái duyệt if blank
    if "Trạng thái duyệt" in df.columns:
        df["Trạng thái duyệt"] = df["Trạng thái duyệt"].replace("", "Chờ duyệt").fillna("Chờ duyệt")
        
    return df

def save_to_gsheets(df, url_or_id, sheet_name):
    client = get_sheets_client()
    if not client:
        raise ConnectionError("Không thể xác thực bằng file credentials.json.")
        
    if "docs.google.com" in url_or_id:
        sh = client.open_by_url(url_or_id)
    else:
        sh = client.open_by_key(url_or_id)
        
    worksheet = sh.worksheet(sheet_name)
    
    # Clean NaN/None for sheet upload
    df_clean = df.fillna("")
    
    # Get current headers
    headers = worksheet.row_values(1)
    
    # Ensure all columns in dataframe exist in headers
    for col in df_clean.columns:
        if col not in headers:
            new_col_idx = len(headers) + 1
            worksheet.update_cell(1, new_col_idx, col)
            headers.append(col)
            
    # Map data correctly to columns
    row_count = len(df_clean)
    cell_list = []
    
    for idx, row in df_clean.iterrows():
        row_num = idx + 2
        for col_name in df_clean.columns:
            col_idx = headers.index(col_name) + 1
            val = str(row[col_name])
            cell_list.append(gspread.Cell(row=row_num, col=col_idx, value=val))
            
    worksheet.update_cells(cell_list)

# --- HEADER SECTION ---
st.markdown("<h1 style='text-align: center; margin-bottom: 0.5rem;'>🎯 AI Lead Scoring & Human Approval</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #64748B; font-size: 1.1rem; margin-bottom: 2rem;'>Hệ thống tự động chấm điểm khách hàng tiềm năng ngành BĐS kết hợp duyệt phê duyệt thủ công</p>", unsafe_allow_html=True)

# --- SIDEBAR CONTROL PANEL ---
with st.sidebar:
    st.markdown("### 🛠️ CẤU HÌNH NGUỒN DỮ LIỆU")
    source = st.radio(
        "Chọn phương thức nạp dữ liệu:",
        ("Dữ liệu mẫu (Thử nghiệm)", "Tải tệp CSV/Excel lên", "Google Sheets (Thời gian thực)")
    )
    
    st.markdown("---")
    
    df_loaded = None
    
    if source == "Dữ liệu mẫu (Thử nghiệm)":
        st.session_state.source_type = "mock"
        if st.button("✨ Nạp dữ liệu mẫu", use_container_width=True):
            st.session_state.df = load_mock_data()
            st.success("Đã nạp 5 leads mẫu để thử nghiệm!")
            
    elif source == "Tải tệp CSV/Excel lên":
        st.session_state.source_type = "file"
        uploaded_file = st.file_uploader("Chọn file CSV hoặc Excel:", type=["csv", "xlsx", "xls"])
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith(".csv"):
                    df_loaded = pd.read_csv(uploaded_file)
                else:
                    df_loaded = pd.read_excel(uploaded_file)
                
                # Check for required columns
                required = ["Họ tên", "Số điện thoại", "Mô tả nhu cầu"]
                missing = [r for r in required if r not in df_loaded.columns]
                
                if missing:
                    st.error(f"Tệp thiếu các cột bắt buộc: {', '.join(missing)}")
                else:
                    for col in ["Điểm số", "Phân loại", "Lý do chấm điểm", "Trạng thái duyệt"]:
                        if col not in df_loaded.columns:
                            df_loaded[col] = ""
                    df_loaded["Trạng thái duyệt"] = df_loaded["Trạng thái duyệt"].replace("", "Chờ duyệt").fillna("Chờ duyệt")
                    st.session_state.df = df_loaded
                    st.success("Tải tệp dữ liệu lên thành công!")
            except Exception as e:
                st.error(f"Lỗi khi đọc tệp: {e}")
                
    elif source == "Google Sheets (Thời gian thực)":
        st.session_state.source_type = "gsheets"
        url = st.text_input("Nhập URL hoặc ID Google Sheet:", value=st.session_state.spreadsheet_url)
        sheet_name = st.text_input("Nhập tên Sheet:", value=st.session_state.sheet_name)
        
        # Guide box for sharing access
        st.info(f"💡 Hãy chia sẻ quyền **Editor** của trang tính cho email này trước:\n`{SERVICE_ACCOUNT_EMAIL}`")
        
        if st.button("📥 Tải dữ liệu từ Google Sheets", use_container_width=True):
            if not url:
                st.error("Vui lòng điền URL hoặc ID Google Sheet.")
            else:
                try:
                    with st.spinner("Đang kết nối & tải dữ liệu..."):
                        loaded_df = fetch_from_gsheets(url, sheet_name)
                        st.session_state.df = loaded_df
                        st.session_state.spreadsheet_url = url
                        st.session_state.sheet_name = sheet_name
                    st.success("Đã tải dữ liệu thành công từ Google Sheets!")
                except Exception as e:
                    st.error(f"Lỗi: {e}")
                    
    st.markdown("---")
    st.markdown("### 🤖 Cấu hình AI Agent")
    st.write("Phương pháp: Chấm điểm Hybrid (Từ khóa + Ngữ cảnh nghiệp vụ).")
    base_score = st.slider("Điểm cơ sở mặc định:", min_value=0, max_value=100, value=50, step=5)

# --- MAIN CONTENT AREA ---
if st.session_state.df is None:
    # Empty State Dashboard
    st.warning("👈 Vui lòng chọn nguồn dữ liệu trong thanh menu bên trái và bấm nạp/tải dữ liệu để bắt đầu.")
    
    # Beautiful presentation of features
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class='metric-card' style='text-align: center;'>
            <h3 style='font-size: 3rem; margin: 0;'>🤖</h3>
            <h4 style='margin: 0.5rem 0;'>AI Scoring tự động</h4>
            <p style='color: #64748B; font-size: 0.9rem;'>Phân tích mô tả nhu cầu khách hàng, tự động phát hiện ngân sách lớn, vị trí đắc địa, hoặc lọc tin rác, liên lạc lỗi.</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class='metric-card' style='text-align: center;'>
            <h3 style='font-size: 3rem; margin: 0;'>✍️</h3>
            <h4 style='margin: 0.5rem 0;'>Human-in-the-Loop</h4>
            <p style='color: #64748B; font-size: 0.9rem;'>Bảng điều khiển cho phép chỉnh sửa điểm số trực tiếp, thay đổi phân loại, và duyệt trạng thái của từng khách hàng nhanh chóng.</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class='metric-card' style='text-align: center;'>
            <h3 style='font-size: 3rem; margin: 0;'>🔄</h3>
            <h4 style='margin: 0.5rem 0;'>Đồng bộ 2 chiều</h4>
            <p style='color: #64748B; font-size: 0.9rem;'>Hỗ trợ xuất file báo cáo hoặc ghi trực tiếp kết quả đã duyệt trở lại trang tính Google Sheets của bạn trong nháy mắt.</p>
        </div>
        """, unsafe_allow_html=True)

else:
    # We have data! Let's display the interactive dashboard.
    df = st.session_state.df
    
    # Ensure numeric types for Score
    df["Điểm số"] = pd.to_numeric(df["Điểm số"], errors='coerce').fillna(base_score).astype(int)
    
    # 1. METRICS DASHBOARD
    total_leads = len(df)
    
    # Dynamic computation of classifications
    vip_leads = len(df[df["Phân loại"] == "VIP / SIÊU TIỀM NĂNG"])
    trash_leads = len(df[df["Phân loại"] == "RÁC / KHÔNG TIỀM NĂNG"])
    med_leads = len(df[df["Phân loại"] == "TRUNG BÌNH"])
    
    # Human approval metrics
    approved_leads = len(df[df["Trạng thái duyệt"].str.contains("Đã duyệt", na=False, case=False)])
    pending_leads = len(df[df["Trạng thái duyệt"] == "Chờ duyệt"])
    
    # Visual grid of metrics
    m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
    with m_col1:
        st.metric(label="Tổng số khách hàng", value=total_leads, delta="Tất cả nguồn")
    with m_col2:
        st.metric(label="⭐ VIP / Siêu tiềm năng", value=vip_leads, delta=f"{vip_leads/total_leads*100:.0f}%" if total_leads > 0 else "0%", delta_color="normal")
    with m_col3:
        st.metric(label="🗑️ Rác / Không tiềm năng", value=trash_leads, delta=f"-{trash_leads/total_leads*100:.0f}%" if total_leads > 0 else "0%", delta_color="inverse")
    with m_col4:
        st.metric(label="💼 Đã duyệt (Con người)", value=approved_leads, delta=f"{approved_leads/total_leads*100:.0f}%" if total_leads > 0 else "0%")
    with m_col5:
        st.metric(label="⏳ Đợi duyệt", value=pending_leads, delta=f"-{pending_leads/total_leads*100:.0f}%" if total_leads > 0 else "0%", delta_color="off")
        
    st.markdown("---")
    
    # 2. CHARTS & ACTIONS
    c_col1, c_col2 = st.columns([2, 5])
    
    with c_col1:
        st.markdown("### 📊 Biểu đồ phân loại")
        chart_data = pd.DataFrame({
            "Phân loại": ["VIP", "Trung bình", "Rác"],
            "Số lượng": [vip_leads, med_leads, trash_leads]
        })
        
        # Altair Pie/Donut Chart
        donut = alt.Chart(chart_data).mark_arc(innerRadius=50).encode(
            theta=alt.Theta(field="Số lượng", type="quantitative"),
            color=alt.Color(field="Phân loại", type="nominal", scale=alt.Scale(
                domain=["VIP", "Trung bình", "Rác"],
                range=["#FF4B4B", "#FFBD45", "#9CA3AF"]
            )),
            tooltip=["Phân loại", "Số lượng"]
        ).properties(width=220, height=220)
        st.altair_chart(donut, use_container_width=True)
        
    with c_col2:
        st.markdown("### 🛠️ Thao tác nhanh")
        
        btn_col1, btn_col2, btn_col3 = st.columns(3)
        
        with btn_col1:
            if st.button("🤖 CHẠY AI SCORING TỰ ĐỘNG", use_container_width=True, type="primary"):
                with st.spinner("AI đang phân tích mô tả nhu cầu khách hàng..."):
                    for idx, row in df.iterrows():
                        demand = row["Mô tả nhu cầu"]
                        score, reason, classification = evaluate_lead(demand)
                        df.at[idx, "Điểm số"] = score
                        df.at[idx, "Phân loại"] = classification
                        df.at[idx, "Lý do chấm điểm"] = reason
                    st.session_state.df = df
                    st.rerun()
                    
        with btn_col2:
            if st.session_state.source_type == "gsheets":
                if st.button("💾 ĐỒNG BỘ LÊN GOOGLE SHEETS", use_container_width=True):
                    try:
                        with st.spinner("Đang lưu dữ liệu lên Google Sheets..."):
                            save_to_gsheets(df, st.session_state.spreadsheet_url, st.session_state.sheet_name)
                            st.success("Đã đồng bộ thành công dữ liệu lên Google Sheets!")
                    except Exception as e:
                        st.error(f"Lỗi đồng bộ: {e}")
            else:
                st.button("💾 ĐỒNG BỘ LÊN GOOGLE SHEETS", use_container_width=True, disabled=True, help="Tính năng chỉ khả dụng khi nguồn nạp dữ liệu là Google Sheets.")
                
        with btn_col3:
            # Download file section
            towrite = io.BytesIO()
            df.to_excel(towrite, index=False, header=True, engine='openpyxl')
            towrite.seek(0)
            st.download_button(
                label="📥 XUẤT BÁO CÁO EXCEL",
                data=towrite,
                file_name="lead_scoring_approved.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    st.markdown("---")
    
    # 3. INTERACTIVE DATA EDITOR (Human-in-the-Loop)
    st.markdown("### ✍️ Bảng chỉnh sửa & Phê duyệt Trạng thái")
    st.caption("Double click vào bất kỳ ô nào để chỉnh sửa điểm, thay đổi phân loại hoặc đặt trạng thái phê duyệt của con người.")
    
    # Streamlit data editor with configuration
    edited_df = st.data_editor(
        df,
        column_config={
            "Họ tên": st.column_config.TextColumn("Họ tên", width="medium"),
            "Số điện thoại": st.column_config.TextColumn("Số điện thoại", width="small"),
            "Mô tả nhu cầu": st.column_config.TextColumn("Mô tả nhu cầu", width="large"),
            "Điểm số": st.column_config.NumberColumn("Điểm số", min_value=0, max_value=100, step=1, format="%d"),
            "Phân loại": st.column_config.SelectboxColumn(
                "Phân loại",
                options=["VIP / SIÊU TIỀM NĂNG", "TRUNG BÌNH", "RÁC / KHÔNG TIỀM NĂNG"],
                width="medium"
            ),
            "Lý do chấm điểm": st.column_config.TextColumn("Lý do chấm điểm", width="medium"),
            "Trạng thái duyệt": st.column_config.SelectboxColumn(
                "Trạng thái duyệt",
                options=["Chờ duyệt", "Đã duyệt (Đúng)", "Đã duyệt (Đã chỉnh sửa)", "Không đồng ý"],
                width="medium"
            )
        },
        num_rows="dynamic",
        use_container_width=True,
        key="lead_editor"
    )
    
    # Real-time state synchronization
    if not edited_df.equals(df):
        st.session_state.df = edited_df
        st.rerun()
