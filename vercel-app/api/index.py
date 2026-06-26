import os
import sys
import io
import json
from pathlib import Path
from typing import List, Dict, Any, Optional

from fastapi import FastAPI, HTTPException, Body
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Thêm thư mục hiện tại và cha vào sys.path để import score_leads
CURRENT_DIR = Path(__file__).parent.resolve()
PARENT_DIR = CURRENT_DIR.parent.resolve()

if str(PARENT_DIR) not in sys.path:
    sys.path.insert(0, str(PARENT_DIR))

try:
    import score_leads
except ImportError:
    # Trường hợp chạy cục bộ hoặc cấu trúc khác
    sys.path.insert(0, str(CURRENT_DIR))
    import score_leads

app = FastAPI(title="AI Lead Scoring API")

# Cấu hình CORS để frontend có thể gọi API dễ dàng
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Models định nghĩa dữ liệu đầu vào
class ScoreRequest(BaseModel):
    url: str

class LeadData(BaseModel):
    id: str
    ten_khach: str
    sdt: str
    nhu_cau_mo_ta: str
    final_score: float
    human_category: str
    reasoning: str
    review_notes: Optional[str] = ""
    __vip_m: Optional[List[str]] = []
    __junk_m: Optional[List[str]] = []

class WriteBackRequest(BaseModel):
    url: str
    leads: List[LeadData]

class ExportRequest(BaseModel):
    leads: List[LeadData]
    integrity_log: List[Dict[str, Any]]
    output_name: Optional[str] = "Lead_Scoring_Report_Final.xlsx"


@app.post("/api/score")
async def api_score(request: ScoreRequest):
    url = request.url.strip()
    if not url:
        raise HTTPException(status_code=400, detail="Đường dẫn Google Sheet không được trống.")
    
    try:
        raw_rows = score_leads.load_data(url)
    except Exception as exc:
        raise HTTPException(
            status_code=500, 
            detail=f"Lỗi tải dữ liệu từ Google Sheet: {str(exc)}. Hãy chắc chắn sheet ở chế độ công khai hoặc đã cấu hình Service Account đúng cách."
        )

    data_list = []
    integrity_log = []
    seen_phones = {}

    for idx, row in enumerate(raw_rows, 2):
        lead_id = str(row.get("id") or f"L{idx-1}").strip()
        ten = str(row.get("ten_khach") or "").strip()
        sdt = str(row.get("sdt") or "").strip()
        nhu_cau = str(row.get("nhu_cau_mo_ta") or "").strip()

        row["id"] = lead_id
        row["ten_khach"] = ten or "Ẩn danh"
        row["sdt"] = sdt or "N/A"
        row["nhu_cau_mo_ta"] = nhu_cau

        # Kiểm tra tính toàn vẹn dữ liệu
        if not ten:
            integrity_log.append({
                "row_num": idx, 
                "id": lead_id, 
                "ten_khach": "N/A",
                "sdt": sdt, 
                "field": "ten_khach",
                "error_msg": "Thiếu tên khách hàng", 
                "action": "Gán mặc định 'Ẩn danh'"
            })
        if not sdt:
            integrity_log.append({
                "row_num": idx, 
                "id": lead_id, 
                "ten_khach": row["ten_khach"],
                "sdt": "Trống", 
                "field": "sdt",
                "error_msg": "Thiếu số điện thoại", 
                "action": "Gán mặc định 'N/A'"
            })
        elif sdt in seen_phones:
            integrity_log.append({
                "row_num": idx, 
                "id": lead_id, 
                "ten_khach": row["ten_khach"],
                "sdt": sdt, 
                "field": "sdt",
                "error_msg": f"SĐT trùng lặp với dòng {seen_phones[sdt]}",
                "action": "Giữ lại và chấm điểm độc lập"
            })
        else:
            seen_phones[sdt] = idx

        if not nhu_cau:
            integrity_log.append({
                "row_num": idx, 
                "id": lead_id, 
                "ten_khach": row["ten_khach"],
                "sdt": row["sdt"], 
                "field": "nhu_cau_mo_ta",
                "error_msg": "Thiếu mô tả nhu cầu", 
                "action": "Không chấm được điểm, giữ nguyên 50"
            })

        ai_score, ai_cat, ai_reason, vip_m, junk_m = score_leads.evaluate_lead(row)

        data_list.append({
            "id": lead_id,
            "ten_khach": row["ten_khach"],
            "sdt": row["sdt"],
            "nhu_cau_mo_ta": nhu_cau,
            "ai_score": ai_score,
            "ai_category": ai_cat,
            "ai_reasoning": ai_reason,
            "final_score": ai_score,
            "human_category": ai_cat,
            "review_notes": "",
            "reasoning": ai_reason,
            "__vip_m": vip_m,
            "__junk_m": junk_m,
        })

    return {
        "success": True,
        "leads": data_list,
        "integrity_log": integrity_log
    }


def build_leads_for_export(leads: List[LeadData]) -> list:
    leads_export = []
    for row in leads:
        score = int(row.final_score)
        cat = row.human_category
        reason = row.reasoning
        vip_m = row.dict().get("__vip_m") or []
        junk_m = row.dict().get("__junk_m") or []

        lead = {
            "id": str(row.id),
            "ten_khach": str(row.ten_khach),
            "sdt": str(row.sdt),
            "nhu_cau_mo_ta": str(row.nhu_cau_mo_ta),
            "__score": (score, cat, reason, vip_m, junk_m),
        }
        leads_export.append(lead)
    return leads_export


@app.post("/api/export")
async def api_export(request: ExportRequest):
    try:
        leads_export = build_leads_for_export(request.leads)
        
        # Sửa lại logs để truyền chính xác vào openpyxl writer
        log = request.integrity_log
        
        buf = io.BytesIO()
        score_leads.create_excel_report(leads_export, log, buf)
        buf.seek(0)
        
        headers = {
            'Content-Disposition': f'attachment; filename="{request.output_name}"'
        }
        return StreamingResponse(
            buf, 
            headers=headers,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Lỗi tạo Excel: {str(exc)}")


@app.post("/api/write-back")
async def api_write_back(request: WriteBackRequest):
    url = request.url.strip()
    if not url:
        raise HTTPException(status_code=400, detail="Đường dẫn Google Sheet không được trống.")
    
    try:
        leads_export = build_leads_for_export(request.leads)
        score_leads.write_scores_to_gspread(leads_export, url)
        return {"success": True, "message": "Cập nhật Google Sheet thành công!"}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Lỗi ghi ngược Google Sheet: {str(exc)}")


# Serve file index.html cho Vercel SPA
@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = PARENT_DIR / "public" / "index.html"
    if not index_path.exists():
        index_path = CURRENT_DIR.parent / "public" / "index.html"
    
    if index_path.exists():
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    
    return HTMLResponse(content="<h1>Frontend index.html not found!</h1>", status_code=404)
