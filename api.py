"""
api.py — Financial Engine FastAPI Backend
รับ request จาก Streamlit / frontend / external systems
"""

import os
import sys
import json
import uuid
import shutil
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# โหลด engine จากโฟลเดอร์เดียวกัน
sys.path.insert(0, str(Path(__file__).parent))

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = FastAPI(
    title="Financial Engine API",
    description="ระบบปิดงบการเงินอัตโนมัติจาก ภ.พ.30",
    version="1.0.0",
)

# CORS — ให้ Streamlit และ frontend เรียกได้
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Directories ───────────────────────────────────────────────
TMP_DIR      = Path("/tmp/financial_engine")
OUTPUT_DIR   = Path("outputs")
TRAINING_DIR = Path("training_data")
for d in [TMP_DIR, OUTPUT_DIR, TRAINING_DIR]:
    d.mkdir(parents=True, exist_ok=True)


# ── Pydantic Models ───────────────────────────────────────────
class ClientConfig(BaseModel):
    company_name:    str   = ""
    tax_id:          str   = ""
    manager_name:    str   = ""
    target_year:     str   = "68"
    retained_begin:  float = 0.0
    admin_expense:   float = 0.0
    inventory_begin: float = 0.0
    inventory_end:   float = 0.0
    other_income:    float = 0.0
    skip_gemini:     bool  = True
    gemini_api_key:  str   = ""

class PrevYearData(BaseModel):
    revenue:        float = 0.0
    other_income:   float = 0.0
    total_revenue:  float = 0.0
    cogs:           float = 0.0
    admin:          float = 0.0
    total_exp:      float = 0.0
    ebit:           float = 0.0
    tax:            float = 0.0
    net_profit:     float = 0.0
    retained_begin: float = 0.0
    retained_end:   float = 0.0

class ExportRequest(BaseModel):
    job_id:         str
    manager_name:   str   = ""
    prev_year_label:str   = "2567"
    prev_year_data: Optional[PrevYearData] = None

class FeedbackRequest(BaseModel):
    job_id:      str
    company:     str
    year:        str
    is_correct:  bool
    note:        str   = ""
    correction:  str   = ""
    metrics:     dict  = {}

# ── In-memory job store (ใช้ Redis ถ้า scale ใหญ่) ────────────
jobs: dict = {}


# ═══════════════════════════════════════════════════════════════
# ENDPOINTS
# ═══════════════════════════════════════════════════════════════

@app.get("/")
def root():
    return {
        "service": "Financial Engine API",
        "version": "1.0.0",
        "status":  "ok",
        "endpoints": [
            "POST /process     — อ่านไฟล์ ภพ.30 และคำนวณ",
            "GET  /job/{id}    — ดูสถานะ job",
            "POST /export      — Export Excel 9 sheets",
            "GET  /download/{id} — Download Excel",
            "POST /feedback    — บันทึก feedback",
            "GET  /stats       — สถิติ training data",
        ]
    }


@app.get("/health")
def health():
    return {"status": "ok", "timestamp": datetime.now().isoformat()}


# ── POST /process ─────────────────────────────────────────────
@app.post("/process")
async def process_file(
    file:            UploadFile = File(...),
    company_name:    str  = Form(""),
    tax_id:          str  = Form(""),
    manager_name:    str  = Form(""),
    target_year:     str  = Form("68"),
    retained_begin:  float = Form(0.0),
    admin_expense:   float = Form(0.0),
    inventory_begin: float = Form(0.0),
    inventory_end:   float = Form(0.0),
    other_income:    float = Form(0.0),
    skip_gemini:     bool  = Form(True),
    gemini_api_key:  str  = Form(""),
):
    """
    อัปโหลดไฟล์ ภพ.30 → ประมวลผล → คืนข้อมูลสรุป
    """
    job_id = str(uuid.uuid4())[:8]
    log.info(f"[{job_id}] รับไฟล์: {file.filename}")

    # บันทึกไฟล์
    input_path = TMP_DIR / f"{job_id}_{file.filename}"
    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # โหลด Engine
    try:
        from financial_engine import FinancialEngine
    except ImportError:
        raise HTTPException(500, "financial_engine.py ไม่พบ")

    api_key = gemini_api_key or os.environ.get("GEMINI_API_KEY", "dummy")

    try:
        engine = FinancialEngine(
            gemini_api_key  = api_key,
            target_year     = target_year,
            retained_begin  = retained_begin,
            admin_expense   = admin_expense,
            other_income    = other_income,
            inventory_begin = inventory_begin,
            inventory_end   = inventory_end,
        )
        summary = engine.process(str(input_path), skip_gemini=skip_gemini)

        # Override ชื่อบริษัท
        if company_name:
            summary.company.name = company_name
        if tax_id:
            summary.company.tax_id = tax_id

        # แก้ total_sale = sale_vat + sale_exempt
        for m in summary.monthly:
            m.total_sale = m.sale_vat + m.sale_exempt

        # คำนวณใหม่
        summary.total_revenue = sum(m.total_sale for m in summary.monthly if m.total_sale > 0)
        summary.total_cogs    = (inventory_begin
                                 + sum(m.purchase for m in summary.monthly)
                                 - inventory_end)
        summary.gross_profit  = summary.total_revenue - summary.total_cogs
        summary.ebit          = summary.gross_profit - admin_expense

        def calc_cit(profit):
            if profit <= 0: return 0.
            tax = 0.
            for limit, rate, prev in [
                (300_000, 0.0, 0),
                (3_000_000, 0.15, 300_000),
                (float('inf'), 0.20, 3_000_000),
            ]:
                if profit > prev:
                    tax += (min(profit, limit) - prev) * rate
            return tax

        cit = calc_cit(max(summary.ebit, 0))
        summary.tax_expense  = -cit
        summary.net_profit   = summary.ebit - cit
        summary.retained_end = retained_begin + summary.net_profit

        # เก็บ job
        jobs[job_id] = {
            "status":    "done",
            "summary":   summary,
            "file":      str(input_path),
            "company":   summary.company.name,
            "year":      target_year,
            "manager":   manager_name,
            "created_at":datetime.now().isoformat(),
        }

        # สร้าง response
        valid = [m for m in summary.monthly if m.total_sale > 0 or m.vat_net != 0]
        monthly_data = [{
            "month":      m.month_th,
            "sale_vat":   round(m.sale_vat, 2),
            "sale_exempt":round(m.sale_exempt, 2),
            "total_sale": round(m.total_sale, 2),
            "purchase":   round(m.purchase, 2),
            "vat_sale":   round(m.vat_sale, 2),
            "vat_buy":    round(m.vat_buy, 2),
            "vat_net":    round(m.vat_net, 2),
            "vat_cumul":  round(m.vat_cumul, 2),
            "confidence": m.confidence,
        } for m in summary.monthly]

        return {
            "job_id":    job_id,
            "status":    "done",
            "company":   summary.company.name,
            "tax_id":    summary.company.tax_id,
            "year":      target_year,
            "summary": {
                "total_revenue": round(summary.total_revenue, 2),
                "total_cogs":    round(summary.total_cogs, 2),
                "gross_profit":  round(summary.gross_profit, 2),
                "gross_margin":  round(summary.gross_profit / summary.total_revenue, 4)
                                 if summary.total_revenue > 0 else 0,
                "admin_expense": round(admin_expense, 2),
                "ebit":          round(summary.ebit, 2),
                "tax_expense":   round(summary.tax_expense, 2),
                "net_profit":    round(summary.net_profit, 2),
                "retained_end":  round(summary.retained_end, 2),
                "total_vat_net": round(summary.total_vat_net, 2),
            },
            "monthly":   monthly_data,
            "anomalies": summary.anomalies,
            "gemini_verdict": summary.gemini_verdict,
        }

    except Exception as e:
        log.error(f"[{job_id}] Error: {e}")
        raise HTTPException(500, f"ประมวลผลไม่สำเร็จ: {str(e)}")


# ── GET /job/{id} ─────────────────────────────────────────────
@app.get("/job/{job_id}")
def get_job(job_id: str):
    """ดูสถานะและผลลัพธ์ของ job"""
    if job_id not in jobs:
        raise HTTPException(404, f"ไม่พบ job: {job_id}")
    job = jobs[job_id]
    return {
        "job_id":     job_id,
        "status":     job["status"],
        "company":    job["company"],
        "year":       job["year"],
        "created_at": job["created_at"],
    }


# ── POST /export ──────────────────────────────────────────────
@app.post("/export")
def export_excel(req: ExportRequest):
    """
    สร้าง Excel 9 sheets จาก job ที่ process แล้ว
    """
    if req.job_id not in jobs:
        raise HTTPException(404, f"ไม่พบ job: {req.job_id}")

    job     = jobs[req.job_id]
    summary = job["summary"]

    try:
        from export_9sheets import build_9sheets
    except ImportError:
        raise HTTPException(500, "export_9sheets.py ไม่พบ")

    output_path = OUTPUT_DIR / f"งบการเงิน_{summary.company.name.replace(' ','_')}_{req.job_id}.xlsx"

    prev_data = req.prev_year_data.dict() if req.prev_year_data else None

    try:
        build_9sheets(
            summary,
            str(output_path),
            manager_name    = req.manager_name or job.get("manager", ""),
            prev_year_label = req.prev_year_label,
            prev_year_data  = prev_data,
        )
        jobs[req.job_id]["output_path"] = str(output_path)
        return {
            "job_id":      req.job_id,
            "status":      "exported",
            "download_url":f"/download/{req.job_id}",
            "filename":    output_path.name,
        }
    except Exception as e:
        raise HTTPException(500, f"Export ไม่สำเร็จ: {str(e)}")


# ── GET /download/{id} ────────────────────────────────────────
@app.get("/download/{job_id}")
def download_excel(job_id: str):
    """Download Excel ที่ export แล้ว"""
    if job_id not in jobs:
        raise HTTPException(404, f"ไม่พบ job: {job_id}")

    output_path = jobs[job_id].get("output_path")
    if not output_path or not Path(output_path).exists():
        raise HTTPException(404, "ยังไม่ได้ export — เรียก POST /export ก่อน")

    return FileResponse(
        path        = output_path,
        filename    = Path(output_path).name,
        media_type  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── POST /feedback ────────────────────────────────────────────
@app.post("/feedback")
def save_feedback(req: FeedbackRequest):
    """
    บันทึก feedback จากนักบัญชี
    ใช้สำหรับ fine-tuning ในอนาคต
    """
    feedback = {
        "timestamp":   datetime.now().isoformat(),
        "job_id":      req.job_id,
        "company":     req.company,
        "year":        req.year,
        "is_correct":  req.is_correct,
        "note":        req.note,
        "correction":  req.correction,
        "metrics":     req.metrics,
    }

    feedback_file = TRAINING_DIR / "feedback_log.jsonl"
    with open(feedback_file, "a", encoding="utf-8") as f:
        f.write(json.dumps(feedback, ensure_ascii=False) + "\n")

    # นับจำนวน feedback ที่มี
    total = sum(1 for _ in open(feedback_file))
    correct = sum(1 for line in open(feedback_file)
                  if json.loads(line).get("is_correct"))

    return {
        "status":       "saved",
        "total_feedback": total,
        "accuracy":     round(correct / total, 2) if total > 0 else 0,
        "ready_for_finetune": total >= 50,
    }


# ── GET /stats ────────────────────────────────────────────────
@app.get("/stats")
def get_stats():
    """สถิติ training data และ job history"""
    feedback_file = TRAINING_DIR / "feedback_log.jsonl"
    total = correct = 0
    companies = {}

    if feedback_file.exists():
        for line in open(feedback_file):
            try:
                fb = json.loads(line)
                total += 1
                if fb.get("is_correct"):
                    correct += 1
                co = fb.get("company","unknown")
                companies[co] = companies.get(co, 0) + 1
            except:
                pass

    return {
        "jobs_in_memory":    len(jobs),
        "total_feedback":    total,
        "correct_feedback":  correct,
        "accuracy":          round(correct / total, 2) if total > 0 else 0,
        "companies":         companies,
        "ready_for_finetune":total >= 50,
    }


# ── GET /list-jobs ────────────────────────────────────────────
@app.get("/list-jobs")
def list_jobs():
    """แสดง jobs ทั้งหมดใน memory"""
    return [{
        "job_id":     jid,
        "company":    j["company"],
        "year":       j["year"],
        "status":     j["status"],
        "created_at": j["created_at"],
        "exported":   "output_path" in j,
    } for jid, j in jobs.items()]


# ── Run ───────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "api:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
    )
