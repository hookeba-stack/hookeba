#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit App for Real Estate Lead Scoring & Human Review
============================================================
Cho phép người dùng tải dữ liệu từ Google Sheets, quét tự động bằng AI,
chỉnh sửa/kiểm duyệt thủ công kết quả phân loại, và xuất báo cáo Excel chuyên nghiệp.
"""

import sys
import os
import io
import re
from datetime import datetime
from pathlib import Path

# Thêm thư mục scripts của skill vào sys.path để import (ưu tiên thư mục cục bộ trong dự án)
PROJECT_ROOT = Path(__file__).parent.resolve()
LOCAL_SKILL_PATH = PROJECT_ROOT / "skills" / "lead-scoring" / "scripts"
if LOCAL_SKILL_PATH.exists():
    sys.path.append(str(LOCAL_SKILL_PATH))
else:
    GLOBAL_SKILL_PATH = Path.home() / ".gemini" / "config" / "skills" / "lead-scoring" / "scripts"
    sys.path.append(str(GLOBAL_SKILL_PATH))

# Tự động cài đặt và import thư viện thiếu
def install_and_import(package, import_name=None):
    if import_name is None:
        import_name = package
    try:
        return __import__(import_name)
    except ImportError:
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            return __import__(import_name)
        except Exception as e:
            print(f"❌ Không thể cài đặt '{package}': {e}")
            return None

streamlit = install_and_import("streamlit")
pandas = install_and_import("pandas")

# Import các logic cốt lõi từ skill
try:
    import score_leads
except ImportError as e:
    streamlit.error(f"Không thể import module 'score_leads' từ thư mục cấu hình: {e}")
    streamlit.info("Đảm bảo rằng skill ckm:lead-scoring đã được khởi tạo chính xác.")
    sys.exit(1)

import streamlit as st
import pandas as pd

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="AI Lead Scoring Center",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS để giao diện trông chuyên nghiệp và sang trọng hơn
st.markdown("""
<style>
    /* Gradient Banner */
    .banner {
        background: linear-gradient(135deg, #1F4E79 0%, #2A5298 100%);
        padding: 25px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    .banner h1 {
        margin: 0;
        font-family: 'Segoe UI', Arial;
        font-weight: 700;
        font-size: 28px;
        letter-spacing: 0.5px;
    }
    .banner p {
        margin: 8px 0 0 0;
        font-size: 14px;
        opacity: 0.9;
        font-style: italic;
    }
    
    /* KPI Card styling */
    .kpi-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .kpi-card-vip {
        border-left: 5px solid #22C55E;
    }
    .kpi-card-normal {
        border-left: 5px solid #3B82F6;
    }
    .kpi-card-junk {
        border-left: 5px solid #EF4444;
    }
    .kpi-val {
        font-size: 24px;
        font-weight: bold;
        color: #1F4E79;
        margin: 5px 0;
    }
    .kpi-lbl {
        font-size: 11px;
        font-weight: bold;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
</style>
""", unsafe_allow_html=True)

# Hiển thị Banner tiêu đề
st.markdown("""
<div class="banner">
    <h1>🎯 TRUNG TÂM KIỂM DUYỆT & CHẤM ĐIỂM KHÁCH HÀNG TIỀM NĂNG</h1>
    <p>AI Lead Scoring System with Human-in-the-Loop Review • Ngành Bất Động Sản</p>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Các hàm trợ giúp
# ---------------------------------------------------------------------------
def convert_df_to_leads(df):
    """Chuyển đổi DataFrame hiện tại thành định dạng list of dicts của score_leads."""
    leads = []
    for _, row in df.iterrows():
        lead_dict = {
            "id": str(row["id"]),
            "ten_khach": str(row["ten_khach"]),
            "sdt": str(row["sdt"]),
            "nhu_cau_mo_ta": str(row["nhu_cau_mo_ta"]),
        }
        
        score = int(row["final_score"])
        cat = row["human_category"]
        reason = row["reasoning"]
        
        # Xác định lại list vip_matched và junk_matched tương ứng
        vip_m = ["Phê duyệt thủ công (VIP)"] if cat == "VIP" else []
        junk_m = ["Phê duyệt thủ công (Rác)"] if cat == "Rác" else []
        
        # Nếu giữ nguyên AI và trùng khớp từ ban đầu
        if cat == row["ai_category"]:
            # Để giữ lại chi tiết từ khóa khớp gốc
            vip_m = row.get("__vip_m", [])
            junk_m = row.get("__junk_m", [])
            
        lead_dict["__score"] = (score, cat, reason, vip_m, junk_m)
        leads.append(lead_dict)
    return leads

def sync_changes(df):
    """
    Đồng bộ và cập nhật điểm số cuối cùng (final_score) cùng lý do (reasoning)
    khi con người thay đổi trạng thái phân loại hoặc ghi chú.
    """
    for idx, row in df.iterrows():
        h_cat = row["human_category"]
        ai_cat = row["ai_category"]
        notes = row["review_notes"]
        
        # 1. Nếu có thay đổi phân loại so với AI
        if h_cat != ai_cat:
            if h_cat == "VIP":
                df.at[idx, "final_score"] = 100
            elif h_cat == "Rác":
                df.at[idx, "final_score"] = 0
            else:
                df.at[idx, "final_score"] = 50
            
            note_str = f" | Ghi chú duyệt: {notes}" if notes else ""
            df.at[idx, "reasoning"] = f"👤 Kiểm duyệt thủ công: Đã chuyển thành {h_cat}{note_str} (AI xếp loại ban đầu: {ai_cat})"
            
        # 2. Nếu giữ nguyên phân loại của AI nhưng có thêm ghi chú kiểm duyệt
        elif notes:
            df.at[idx, "final_score"] = row["ai_score"]
            df.at[idx, "reasoning"] = f"👤 Kiểm duyệt thủ công: Giữ nguyên {h_cat} | Ghi chú duyệt: {notes}"
            
        # 3. Không sửa gì và không ghi chú (Trở về trạng thái AI thuần túy)
        else:
            df.at[idx, "final_score"] = row["ai_score"]
            df.at[idx, "reasoning"] = row["ai_reasoning"]
            
    return df

# ---------------------------------------------------------------------------
# Sidebar cấu hình đầu vào
# ---------------------------------------------------------------------------
st.sidebar.header("⚙️ Cấu Hình Đầu Vào")

default_url = "https://docs.google.com/spreadsheets/d/1EaHaNMNmqz2Yy-3DpaNktbi4Ii0vRvaz4lwjpL71zYg/edit?usp=sharing"
url_input = st.sidebar.text_input("🔗 Link Google Sheet hoặc tệp cục bộ:", value=default_url)

output_filename = st.sidebar.text_input("💾 Tên file báo cáo xuất ra:", value="Lead_Scoring_Report_Final.xlsx")

st.sidebar.markdown("---")

run_button = st.sidebar.button("⚡ Tải Dữ Liệu & Quét AI", use_container_width=True)

# ---------------------------------------------------------------------------
# Xử lý Logic Tải và Quét AI
# ---------------------------------------------------------------------------
if run_button:
    with st.spinner("⏳ Đang tải dữ liệu và phân tích chất lượng bằng AI..."):
        try:
            # Tải dữ liệu bằng hàm load_data của score_leads
            raw_rows = score_leads.load_data(url_input)
            
            # Quét và chấm điểm từng dòng
            scored_leads = []
            integrity_log = []
            seen_phones = {}
            
            for idx, row in enumerate(raw_rows, 2):
                lead_id = str(row.get("id", "")).strip() if row.get("id") else f"L{idx-1}"
                ten = str(row.get("ten_khach", "")).strip() if row.get("ten_khach") else ""
                sdt = str(row.get("sdt", "")).strip() if row.get("sdt") else ""
                nhu_cau = str(row.get("nhu_cau_mo_ta", "")).strip() if row.get("nhu_cau_mo_ta") else ""
                
                row["id"] = lead_id
                row["ten_khach"] = ten if ten else "Ẩn danh"
                row["sdt"] = sdt if sdt else "N/A"
                row["nhu_cau_mo_ta"] = nhu_cau
                
                # Kiểm tra tính toàn vẹn dữ liệu
                if not ten:
                    integrity_log.append({
                        "row_num": idx, "id": lead_id, "ten_khach": "N/A", "sdt": sdt,
                        "field": "ten_khach", "error_msg": "Thiếu tên khách hàng", "action": "Gán mặc định 'Ẩn danh'"
                    })
                if not sdt or sdt == "N/A":
                    integrity_log.append({
                        "row_num": idx, "id": lead_id, "ten_khach": row["ten_khach"], "sdt": "Trống",
                        "field": "sdt", "error_msg": "Thiếu số điện thoại liên lạc", "action": "Gán mặc định 'N/A'"
                    })
                else:
                    if sdt in seen_phones:
                        integrity_log.append({
                            "row_num": idx, "id": lead_id, "ten_khach": row["ten_khach"], "sdt": sdt,
                            "field": "sdt", "error_msg": f"Số điện thoại trùng lặp với dòng {seen_phones[sdt]}", "action": "Giữ lại và chấm điểm độc lập"
                        })
                    else:
                        seen_phones[sdt] = idx
                        
                if not nhu_cau:
                    integrity_log.append({
                        "row_num": idx, "id": lead_id, "ten_khach": row["ten_khach"], "sdt": row["sdt"],
                        "field": "nhu_cau_mo_ta", "error_msg": "Thiếu mô tả nhu cầu bất động sản", "action": "Không thể chấm điểm, giữ điểm 50"
                    })
                
                # Đánh giá bằng AI
                score, cat, reason, vip_m, junk_m = score_leads.evaluate_lead(row)
                row["__score"] = (score, cat, reason, vip_m, junk_m)
                scored_leads.append(row)
                
            # Tạo DataFrame để lưu trữ vào session state
            data_list = []
            for item in scored_leads:
                ai_score, ai_cat, ai_reason, vip_m, junk_m = item["__score"]
                data_list.append({
                    "id": item["id"],
                    "ten_khach": item["ten_khach"],
                    "sdt": item["sdt"],
                    "nhu_cau_mo_ta": item["nhu_cau_mo_ta"],
                    "ai_score": ai_score,
                    "ai_category": ai_cat,
                    "ai_reasoning": ai_reason,
                    "final_score": ai_score,
                    "human_category": ai_cat,
                    "review_notes": "",
                    "reasoning": ai_reason,
                    "__vip_m": vip_m,
                    "__junk_m": junk_m
                })
                
            st.session_state.df = pd.DataFrame(data_list)
            st.session_state.integrity_log = integrity_log
            st.session_state.original_url = url_input
            st.toast("✅ Đã tải dữ liệu và quét AI thành công!", icon="🚀")
            
        except Exception as e:
            st.error(f"❌ Lỗi khi tải dữ liệu: {e}")

# ---------------------------------------------------------------------------
# Giao Diện Chính Khi Có Dữ Liệu
# ---------------------------------------------------------------------------
if "df" in st.session_state:
    df = st.session_state.df
    integrity_log = st.session_state.integrity_log
    
    # 1. Tính toán các chỉ số Dashboard động
    total_leads = len(df)
    vip_leads = sum(df["human_category"] == "VIP")
    normal_leads = sum(df["human_category"] == "Bình thường")
    junk_leads = sum(df["human_category"] == "Rác")
    avg_score = df["final_score"].mean()
    
    # Hiển thị KPI Blocks
    kpi_cols = st.columns(5)
    with kpi_cols[0]:
        st.markdown(f'<div class="kpi-card"><div class="kpi-lbl">Tổng số khách</div><div class="kpi-val">{total_leads}</div></div>', unsafe_allow_html=True)
    with kpi_cols[1]:
        st.markdown(f'<div class="kpi-card kpi-card-vip"><div class="kpi-lbl">Khách VIP ⭐</div><div class="kpi-val" style="color:#22C55E;">{vip_leads}</div></div>', unsafe_allow_html=True)
    with kpi_cols[2]:
        st.markdown(f'<div class="kpi-card kpi-card-normal"><div class="kpi-lbl">Bình thường 👥</div><div class="kpi-val" style="color:#3B82F6;">{normal_leads}</div></div>', unsafe_allow_html=True)
    with kpi_cols[3]:
        st.markdown(f'<div class="kpi-card kpi-card-junk"><div class="kpi-lbl">Khách Rác 🗑️</div><div class="kpi-val" style="color:#EF4444;">{junk_leads}</div></div>', unsafe_allow_html=True)
    with kpi_cols[4]:
        st.markdown(f'<div class="kpi-card"><div class="kpi-lbl">Điểm trung bình</div><div class="kpi-val">{avg_score:.1f}</div></div>', unsafe_allow_html=True)
        
    st.markdown("---")
    
    # 2. Bảng kiểm duyệt của con người
    st.subheader("📝 Bảng Kiểm Duyệt Khách Hàng (Human-in-the-loop)")
    st.info("💡 Bạn có thể click đúp vào cột **Phân Loại Kiểm Duyệt** để thay đổi phân loại, hoặc điền cột **Ghi Chú Kiểm Duyệt** để ghi lại lý do kiểm duyệt.")
    
    # Định nghĩa cấu hình hiển thị cột cho st.data_editor
    column_config = {
        "id": st.column_config.TextColumn("Mã Lead", width="small", disabled=True),
        "ten_khach": st.column_config.TextColumn("Tên Khách Hàng", width="medium", disabled=True),
        "sdt": st.column_config.TextColumn("Số Điện Thoại", width="small", disabled=True),
        "nhu_cau_mo_ta": st.column_config.TextColumn("Mô Tả Nhu Cầu", width="large", disabled=True),
        "ai_score": st.column_config.NumberColumn("Điểm AI", width="small", format="%d", disabled=True),
        "ai_category": st.column_config.TextColumn("Phân Loại AI", width="small", disabled=True),
        "final_score": st.column_config.NumberColumn("Điểm Cuối Cùng", width="small", format="%d", disabled=True),
        "human_category": st.column_config.SelectboxColumn(
            "Phân Loại Kiểm Duyệt",
            options=["VIP", "Bình thường", "Rác"],
            width="medium",
            required=True
        ),
        "review_notes": st.column_config.TextColumn("Ghi Chú Kiểm Duyệt", width="medium"),
        "reasoning": st.column_config.TextColumn("Giải Trình Chi Tiết", width="large", disabled=True)
    }
    
    # Loại bỏ các cột phụ trợ ra khỏi giao diện editor
    display_df = df.drop(columns=["ai_reasoning", "__vip_m", "__junk_m"])
    
    # Hiển thị data_editor để người dùng tương tác
    edited_df = st.data_editor(
        display_df,
        column_config=column_config,
        use_container_width=True,
        num_rows="fixed",
        key="data_editor_widget"
    )
    
    # Kiểm tra xem có thay đổi nào giữa dữ liệu hiển thị gốc và dữ liệu sau edit không
    # Nếu có thay đổi, đồng bộ lại điểm và lý do, sau đó lưu lại vào session state
    if not edited_df.equals(display_df):
        synced_df = sync_changes(edited_df)
        # Ghi nhận lại các cột bổ trợ ẩn
        synced_df["ai_reasoning"] = df["ai_reasoning"]
        synced_df["__vip_m"] = df["__vip_m"]
        synced_df["__junk_m"] = df["__junk_m"]
        
        st.session_state.df = synced_df
        st.rerun()
        
    st.markdown("---")
    
    # 3. Hành động xuất dữ liệu (Export Options)
    st.subheader("💾 Hành Động Xuất Kết Quả")
    
    action_cols = st.columns(3)
    
    # Tùy chọn A: Tải xuống báo cáo Excel trực tiếp qua browser
    with action_cols[0]:
        st.write("📊 **Tải báo cáo Excel (Dashboard + Bảng điểm)**")
        leads_for_excel = convert_df_to_leads(df)
        
        # Viết Excel vào BytesIO buffer
        buffer = io.BytesIO()
        try:
            score_leads.create_excel_report(leads_for_excel, integrity_log, buffer)
            buffer.seek(0)
            
            st.download_button(
                label="📥 Tải Xuống File Báo Cáo Excel",
                data=buffer,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as excel_err:
            st.error(f"Lỗi tạo Excel: {excel_err}")
            
    # Tùy chọn B: Lưu file Excel cục bộ vào Workspace
    with action_cols[1]:
        st.write("📂 **Lưu báo cáo vào thư mục dự án cục bộ**")
        workspace_save_path = Path(__file__).parent.resolve() / output_filename
        if st.button("💾 Lưu File Cục Bộ Vào Workspace", use_container_width=True):
            try:
                leads_for_excel = convert_df_to_leads(df)
                score_leads.create_excel_report(leads_for_excel, integrity_log, workspace_save_path)
                st.success(f"Đã lưu thành công tại: {workspace_save_path}")
            except Exception as save_err:
                st.error(f"Lỗi khi lưu file: {save_err}")
                
    # Tùy chọn C: Ghi ngược lại Google Sheet gốc (nếu có tài khoản)
    with action_cols[2]:
        st.write("🔄 **Ghi đè kết quả trực tiếp lên Google Sheet**")
        if st.button("📝 Cập Nhật Ngược Google Sheet", use_container_width=True):
            if st.session_state.original_url.startswith("http"):
                with st.spinner("🔄 Đang gửi dữ liệu cập nhật..."):
                    try:
                        leads_for_excel = convert_df_to_leads(df)
                        score_leads.write_scores_to_gspread(leads_for_excel, st.session_state.original_url)
                        st.success("Đã hoàn tất cập nhật lên Google Sheet!")
                    except Exception as sheet_err:
                        st.error(f"Cập nhật thất bại: {sheet_err}")
            else:
                st.warning("Không thể cập nhật vì dữ liệu đầu vào không phải Google Sheet online.")
                
    # 4. Hiển thị Nhật ký Toàn vẹn Dữ liệu (nếu có lỗi)
    if integrity_log:
        st.markdown("---")
        with st.expander(f"⚠️ Nhật Ký Cảnh Báo Toàn Vẹn Dữ Liệu ({len(integrity_log)} lỗi)"):
            log_df = pd.DataFrame(integrity_log)
            st.dataframe(
                log_df.rename(columns={
                    "row_num": "Dòng", "id": "Mã Lead", "ten_khach": "Tên Khách Hàng",
                    "sdt": "Số Điện Thoại", "field": "Cột Bị Lỗi", "error_msg": "Mô Tả Lỗi", "action": "Cách Xử Lý"
                }),
                use_container_width=True,
                hide_index=True
            )
else:
    # Trường hợp chưa tải dữ liệu
    st.info("👉 Vui lòng nhập link dữ liệu hoặc giữ mặc định ở Sidebar và nhấn nút **Tải Dữ Liệu & Quét AI** để bắt đầu phân tích.")
