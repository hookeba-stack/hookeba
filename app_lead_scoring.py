#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit App - AI Lead Scoring Center
Real Estate CRM | Human-in-the-Loop Review System
"""

import sys
import io
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd

# Import core logic from local score_leads.py (same directory)
PROJECT_ROOT = Path(__file__).parent.resolve()
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

try:
    import score_leads
except ImportError as exc:
    st.error(f"Cannot import score_leads module: {exc}")
    st.stop()

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Trung tâm Chấm điểm Lead AI",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.banner {
    background: linear-gradient(135deg, #0f2044 0%, #1a3a6b 50%, #1F4E79 100%);
    padding: 30px 36px;
    border-radius: 16px;
    color: white;
    margin-bottom: 28px;
    box-shadow: 0 8px 32px rgba(31,78,121,0.3);
    border: 1px solid rgba(255,255,255,0.08);
}
.banner h1 {
    margin: 0 0 6px 0;
    font-weight: 800;
    font-size: 26px;
    letter-spacing: 0.3px;
}
.banner p {
    margin: 0;
    font-size: 14px;
    opacity: 0.75;
    font-weight: 400;
}

.kpi-grid { display: flex; gap: 14px; margin-bottom: 24px; }
.kpi-card {
    flex: 1;
    background: #ffffff;
    border-radius: 12px;
    padding: 18px 20px;
    text-align: center;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    border: 1px solid #E8EDF5;
    transition: transform 0.2s, box-shadow 0.2s;
}
.kpi-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.1); }
.kpi-card.vip   { border-top: 4px solid #22C55E; }
.kpi-card.norm  { border-top: 4px solid #3B82F6; }
.kpi-card.junk  { border-top: 4px solid #EF4444; }
.kpi-card.total { border-top: 4px solid #1F4E79; }
.kpi-card.avg   { border-top: 4px solid #F59E0B; }
.kpi-val { font-size: 32px; font-weight: 800; color: #0f2044; margin: 6px 0 4px; line-height: 1; }
.kpi-lbl { font-size: 10px; font-weight: 700; color: #6B7280; text-transform: uppercase; letter-spacing: 0.8px; }
.kpi-icon { font-size: 20px; }

.badge-vip  { background:#E2EFDA; color:#375623; padding:3px 10px; border-radius:20px; font-weight:700; font-size:12px; }
.badge-junk { background:#FCE4D6; color:#C65911; padding:3px 10px; border-radius:20px; font-weight:700; font-size:12px; }
.badge-norm { background:#DDEBF7; color:#1F4E79; padding:3px 10px; border-radius:20px; font-weight:700; font-size:12px; }

.stButton>button {
    border-radius: 10px;
    font-weight: 600;
    transition: all 0.2s;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Header Banner
# ---------------------------------------------------------------------------
st.markdown("""
<div class="banner">
    <h1>🎯 TRUNG TÂM CHẤM ĐIỂM LEAD AI</h1>
    <p>Hệ thống chấm điểm khách hàng tiềm năng bất động sản • Kết hợp AI + Kiểm duyệt thủ công</p>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
CATEGORY_MAP = {
    "VIP":          "VIP",
    "Bình thường":  "Bình thường",
    "Rác":          "Rác",
}

def sync_changes(df: pd.DataFrame) -> pd.DataFrame:
    for idx, row in df.iterrows():
        h_cat  = row["human_category"]
        ai_cat = row["ai_category"]
        notes  = row.get("review_notes", "")

        if h_cat != ai_cat:
            if h_cat == "VIP":
                df.at[idx, "final_score"] = 100
            elif h_cat == "Rác":
                df.at[idx, "final_score"] = 0
            else:
                df.at[idx, "final_score"] = 50
            note_str = f" | Ghi chú: {notes}" if notes else ""
            df.at[idx, "reasoning"] = (
                f"[Thủ công] Chuyển thành {h_cat}{note_str} "
                f"(AI: {ai_cat})"
            )
        elif notes:
            df.at[idx, "final_score"] = row["ai_score"]
            df.at[idx, "reasoning"] = (
                f"[Thủ công] Giữ nguyên {h_cat} | Ghi chú: {notes}"
            )
        else:
            df.at[idx, "final_score"] = row["ai_score"]
            df.at[idx, "reasoning"]   = row["ai_reasoning"]
    return df


def build_leads_for_export(df: pd.DataFrame) -> list:
    leads = []
    for _, row in df.iterrows():
        score = int(row["final_score"])
        cat   = row["human_category"]
        reason = row["reasoning"]
        vip_m  = row.get("__vip_m",  [])
        junk_m = row.get("__junk_m", [])

        lead = {
            "id":           str(row["id"]),
            "ten_khach":    str(row["ten_khach"]),
            "sdt":          str(row["sdt"]),
            "nhu_cau_mo_ta": str(row["nhu_cau_mo_ta"]),
            "__score":      (score, cat, reason, vip_m, junk_m),
        }
        leads.append(lead)
    return leads


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Cấu Hình")

    default_url = (
        "https://docs.google.com/spreadsheets/d/"
        "1EaHaNMNmqz2Yy-3DpaNktbi4Ii0vRvaz4lwjpL71zYg/edit?usp=sharing"
    )
    url_input = st.text_input("🔗 Đường dẫn Google Sheet / File:", value=default_url)
    output_name = st.text_input("📄 Tên file xuất:", value="Lead_Scoring_Report_Final.xlsx")

    st.divider()
    run_btn = st.button("⚡ Tải Dữ Liệu & Quét AI", use_container_width=True, type="primary")

    st.divider()
    st.caption("Phiên bản 2.0 | AI Lead Scoring")

# ---------------------------------------------------------------------------
# Main logic: Load & Score
# ---------------------------------------------------------------------------
if run_btn:
    with st.spinner("Đang tải dữ liệu và phân tích bằng AI..."):
        try:
            raw_rows = score_leads.load_data(url_input)
        except Exception as exc:
            st.error(f"Lỗi tải dữ liệu: {exc}")
            st.stop()

        data_list    = []
        integrity_log = []
        seen_phones  = {}

        for idx, row in enumerate(raw_rows, 2):
            lead_id = str(row.get("id") or f"L{idx-1}").strip()
            ten     = str(row.get("ten_khach") or "").strip()
            sdt     = str(row.get("sdt") or "").strip()
            nhu_cau = str(row.get("nhu_cau_mo_ta") or "").strip()

            row["id"]          = lead_id
            row["ten_khach"]   = ten or "Ẩn danh"
            row["sdt"]         = sdt or "N/A"
            row["nhu_cau_mo_ta"] = nhu_cau

            # Integrity checks
            if not ten:
                integrity_log.append({"row_num": idx, "id": lead_id, "ten_khach": "N/A",
                    "sdt": sdt, "field": "ten_khach",
                    "error_msg": "Thiếu tên khách hàng", "action": "Gán mặc định 'Ẩn danh'"})
            if not sdt:
                integrity_log.append({"row_num": idx, "id": lead_id, "ten_khach": row["ten_khach"],
                    "sdt": "Trống", "field": "sdt",
                    "error_msg": "Thiếu số điện thoại", "action": "Gán mặc định 'N/A'"})
            elif sdt in seen_phones:
                integrity_log.append({"row_num": idx, "id": lead_id, "ten_khach": row["ten_khach"],
                    "sdt": sdt, "field": "sdt",
                    "error_msg": f"SĐT trùng lặp với dòng {seen_phones[sdt]}",
                    "action": "Giữ lại và chấm điểm độc lập"})
            else:
                seen_phones[sdt] = idx

            if not nhu_cau:
                integrity_log.append({"row_num": idx, "id": lead_id, "ten_khach": row["ten_khach"],
                    "sdt": row["sdt"], "field": "nhu_cau_mo_ta",
                    "error_msg": "Thiếu mô tả nhu cầu", "action": "Không chấm được điểm, giữ nguyên 50"})

            ai_score, ai_cat, ai_reason, vip_m, junk_m = score_leads.evaluate_lead(row)

            data_list.append({
                "id":           lead_id,
                "ten_khach":    row["ten_khach"],
                "sdt":          row["sdt"],
                "nhu_cau_mo_ta": nhu_cau,
                "ai_score":     ai_score,
                "ai_category":  ai_cat,
                "ai_reasoning": ai_reason,
                "final_score":  ai_score,
                "human_category": ai_cat,
                "review_notes": "",
                "reasoning":    ai_reason,
                "__vip_m":      vip_m,
                "__junk_m":     junk_m,
            })

        st.session_state.df            = pd.DataFrame(data_list)
        st.session_state.integrity_log = integrity_log
        st.session_state.source_url    = url_input
        st.toast(f"Đã tải và phân tích thành công {len(data_list)} khách hàng tiềm năng!", icon="✅")

# ---------------------------------------------------------------------------
# Display section (when data loaded)
# ---------------------------------------------------------------------------
if "df" in st.session_state:
    df  = st.session_state.df
    log = st.session_state.integrity_log

    total    = len(df)
    n_vip    = (df["human_category"] == "VIP").sum()
    n_norm   = (df["human_category"] == "Bình thường").sum()
    n_junk   = (df["human_category"] == "Rác").sum()
    avg_sc   = df["final_score"].mean()

    # KPI cards
    kpi_html = f"""
    <div class="kpi-grid">
        <div class="kpi-card total">
            <div class="kpi-icon">👥</div>
            <div class="kpi-val">{total}</div>
            <div class="kpi-lbl">Tổng Số Khách</div>
        </div>
        <div class="kpi-card vip">
            <div class="kpi-icon">⭐</div>
            <div class="kpi-val" style="color:#16A34A">{n_vip}</div>
            <div class="kpi-lbl">Khách VIP</div>
        </div>
        <div class="kpi-card norm">
            <div class="kpi-icon">👤</div>
            <div class="kpi-val" style="color:#2563EB">{n_norm}</div>
            <div class="kpi-lbl">Bình Thường</div>
        </div>
        <div class="kpi-card junk">
            <div class="kpi-icon">🗑️</div>
            <div class="kpi-val" style="color:#DC2626">{n_junk}</div>
            <div class="kpi-lbl">Khách Rác</div>
        </div>
        <div class="kpi-card avg">
            <div class="kpi-icon">📊</div>
            <div class="kpi-val" style="color:#D97706">{avg_sc:.1f}</div>
            <div class="kpi-lbl">Điểm TB</div>
        </div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    st.divider()

    # Review table
    st.subheader("📝 Bảng Kiểm Duyệt Khách Hàng")
    st.info(
        "💡 Click đúp vào ô ở cột **Phân Loại KD** để thay đổi phân loại của khách hàng. "
        "Điền thông tin vào cột **Ghi Chú** để ghi nhận lý do kiểm duyệt thủ công."
    )

    col_cfg = {
        "id":             st.column_config.TextColumn("Mã Lead",       width="small",  disabled=True),
        "ten_khach":      st.column_config.TextColumn("Tên Khách Hàng", width="medium", disabled=True),
        "sdt":            st.column_config.TextColumn("Số ĐT",          width="small",  disabled=True),
        "nhu_cau_mo_ta":  st.column_config.TextColumn("Mô Tả Nhu Cầu",  width="large",  disabled=True),
        "ai_score":       st.column_config.NumberColumn("Điểm AI",      width="small",  format="%d", disabled=True),
        "ai_category":    st.column_config.TextColumn("Phân Loại AI",   width="small",  disabled=True),
        "final_score":    st.column_config.NumberColumn("Điểm Cuối",    width="small",  format="%d", disabled=True),
        "human_category": st.column_config.SelectboxColumn(
            "Phân Loại KD",
            options=["VIP", "Bình thường", "Rác"],
            width="medium",
            required=True,
        ),
        "review_notes":   st.column_config.TextColumn("Ghi Chú",       width="medium"),
        "reasoning":      st.column_config.TextColumn("Giải Trình",     width="large",  disabled=True),
    }

    display_df = df.drop(columns=["ai_reasoning", "__vip_m", "__junk_m"], errors="ignore")

    edited_df = st.data_editor(
        display_df,
        column_config=col_cfg,
        use_container_width=True,
        num_rows="fixed",
        key="editor_v2",
    )

    if not edited_df.equals(display_df):
        synced = sync_changes(edited_df.copy())
        synced["ai_reasoning"] = df["ai_reasoning"]
        synced["__vip_m"]      = df["__vip_m"]
        synced["__junk_m"]     = df["__junk_m"]
        st.session_state.df = synced
        st.rerun()

    st.divider()

    # Export options
    st.subheader("💾 Xuất Kết Quả")
    exp_col1, exp_col2, exp_col3 = st.columns(3)

    with exp_col1:
        st.write("**📊 Tải Báo Cáo Excel**")
        buf = io.BytesIO()
        try:
            leads_export = build_leads_for_export(df)
            score_leads.create_excel_report(leads_export, log, buf)
            buf.seek(0)
            st.download_button(
                label="📥 Tải Xuống Excel",
                data=buf,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Lỗi tạo Excel: {exc}")

    with exp_col2:
        st.write("**📂 Lưu File Cục Bộ**")
        save_path = PROJECT_ROOT / output_name
        if st.button("💾 Lưu Vào Thư Mục Dự Án", use_container_width=True):
            try:
                leads_export = build_leads_for_export(df)
                score_leads.create_excel_report(leads_export, log, save_path)
                st.success(f"Đã lưu: {save_path}")
            except Exception as exc:
                st.error(f"Lỗi: {exc}")

    with exp_col3:
        st.write("**🔄 Ghi Ngược Google Sheet**")
        if st.button("📝 Cập Nhật Google Sheet", use_container_width=True):
            src = st.session_state.get("source_url", "")
            if src.startswith("http"):
                with st.spinner("Đang cập nhật..."):
                    try:
                        leads_export = build_leads_for_export(df)
                        score_leads.write_scores_to_gspread(leads_export, src)
                        st.success("Cập nhật Google Sheet thành công!")
                    except Exception as exc:
                        st.error(f"Thất bại: {exc}")
            else:
                st.warning("Nguồn dữ liệu không phải Google Sheet.")

    # Integrity log
    if log:
        st.divider()
        with st.expander(f"⚠️ Nhật Ký Cảnh Báo Dữ Liệu ({len(log)} cảnh báo)"):
            st.dataframe(
                pd.DataFrame(log).rename(columns={
                    "row_num":   "Dòng",
                    "id":        "Mã Lead",
                    "ten_khach": "Tên Khách",
                    "sdt":       "Số ĐT",
                    "field":     "Trường Lỗi",
                    "error_msg": "Mô Tả Lỗi",
                    "action":    "Xử Lý",
                }),
                use_container_width=True,
                hide_index=True,
            )

else:
    st.info(
        "👉 Nhập link Google Sheet vào Sidebar (hoặc giữ mặc định) "
        "và nhấn **Tải Dữ Liệu & Quét AI** để bắt đầu."
    )
