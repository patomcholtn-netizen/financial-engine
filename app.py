"""
app.py — Financial Engine Web UI
Streamlit app สำหรับทีม 2-5 คน
"""
import streamlit as st
import sys, os, json, time, shutil
from pathlib import Path
from datetime import datetime

# ── Page config ─────────────────────────────────────────────
st.set_page_config(
    page_title="Financial Engine",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ───────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Sarabun', sans-serif !important; }

.main-header {
    background: linear-gradient(135deg, #1E3A5F 0%, #1A56DB 100%);
    padding: 2rem 2.5rem; border-radius: 16px; margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(26,86,219,0.25);
}
.main-header h1 { color: #fff; font-size: 2rem; font-weight: 700; margin: 0; }
.main-header p  { color: #BFD9F5; margin: 0.25rem 0 0; font-size: 1rem; }

.metric-card {
    background: #F8FAFF; border: 1px solid #DBEAFE;
    border-radius: 12px; padding: 1.25rem 1.5rem; text-align: center;
}
.metric-card .label { color: #6B7280; font-size: 0.85rem; margin-bottom: 0.25rem; }
.metric-card .value { color: #1E3A5F; font-size: 1.5rem; font-weight: 700; }
.metric-card .value.green { color: #065F46; }
.metric-card .value.red   { color: #DC2626; }

.anomaly-card {
    background: #FEF3C7; border-left: 4px solid #F59E0B;
    border-radius: 8px; padding: 0.75rem 1rem; margin: 0.5rem 0;
}
.anomaly-card.high { background: #FEF2F2; border-color: #DC2626; }

.step-badge {
    background: #1A56DB; color: white; border-radius: 50%;
    width: 28px; height: 28px; display: inline-flex;
    align-items: center; justify-content: center;
    font-weight: 700; font-size: 0.85rem; margin-right: 0.5rem;
}

.status-ok   { color: #065F46; font-weight: 600; }
.status-warn { color: #D97706; font-weight: 600; }
.status-err  { color: #DC2626; font-weight: 600; }

div[data-testid="stExpander"] { border: 1px solid #DBEAFE; border-radius: 10px; }
div.stButton > button {
    background: #1A56DB; color: white; border: none;
    border-radius: 8px; padding: 0.6rem 2rem; font-weight: 600;
    font-family: 'Sarabun', sans-serif; font-size: 1rem;
    transition: all 0.2s;
}
div.stButton > button:hover { background: #1E3A5F; transform: translateY(-1px); }
</style>
""", unsafe_allow_html=True)


# ── Header ───────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 Financial Engine</h1>
    <p>ระบบปิดงบการเงินอัตโนมัติจาก ภ.พ.30</p>
</div>
""", unsafe_allow_html=True)


# ── Sidebar: ข้อมูลลูกค้า ────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏢 ข้อมูลลูกค้า")
    st.markdown("---")

    company_name = st.text_input(
        "ชื่อบริษัท",
        value="ห้างหุ้นส่วนจำกัด โรจน์เจริญ กันทรลักษ์",
        help="ชื่อเต็มของห้างหรือบริษัท"
    )
    tax_id = st.text_input("เลขผู้เสียภาษี", value="0333561000131")
    manager_name = st.text_input("ชื่อผู้จัดการ/หุ้นส่วน", value="นางสาวปราณี พูลสุข")
    target_year = st.selectbox("ปี พ.ศ. (2 หลัก)", ["68","67","69"], index=0)

    st.markdown("---")
    st.markdown("### 💰 ตัวเลขเพิ่มเติม")

    retained_begin = st.number_input(
        "กำไรสะสมต้นปี (บาท)",
        value=2278127.06, format="%.2f"
    )
    admin_expense = st.number_input(
        "ค่าใช้จ่ายบริหาร (บาท)",
        value=0.00, format="%.2f",
        help="ค่าทำบัญชี ค่าสอบบัญชี เงินเดือน ฯลฯ"
    )
    inventory_begin = st.number_input(
        "สินค้าต้นงวด (บาท)",
        value=409280.74, format="%.2f"
    )
    inventory_end = st.number_input(
        "สินค้าปลายงวด (บาท)",
        value=409280.74, format="%.2f"
    )

    st.markdown("---")
    st.markdown("### 🤖 AI Settings")
    use_gemini = st.toggle("ใช้ Gemini วิเคราะห์", value=False)
    if use_gemini:
        gemini_key = st.text_input(
            "Gemini API Key", type="password",
            help="หรือตั้งค่า GEMINI_API_KEY ใน environment"
        )
    else:
        gemini_key = os.environ.get("GEMINI_API_KEY", "dummy")


# ── Main: Upload ─────────────────────────────────────────────
col_steps = st.columns([1,1,1])
for i, (icon, label) in enumerate([
    ("1️⃣", "อัปโหลดไฟล์"),
    ("2️⃣", "ตรวจสอบข้อมูล"),
    ("3️⃣", "Export Excel"),
]):
    with col_steps[i]:
        bg = "#1E3A5F" if st.session_state.get(f"step_{i+1}") else "#F3F4F6"
        fg = "#fff" if st.session_state.get(f"step_{i+1}") else "#6B7280"
        st.markdown(f"""
        <div style="background:{bg};color:{fg};border-radius:10px;
             padding:0.75rem;text-align:center;font-weight:600;">
            {icon} {label}
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "📂 อัปโหลดไฟล์ ภ.พ.30 (.xlsx หรือ .xls)",
    type=["xlsx","xls"],
    help="ลากไฟล์มาวางหรือคลิกเพื่อเลือก"
)

if not uploaded_file:
    st.info("👆 อัปโหลดไฟล์ ภ.พ.30 เพื่อเริ่มต้น")
    st.stop()

st.session_state["step_1"] = True


# ── Process ──────────────────────────────────────────────────
# บันทึกไฟล์ชั่วคราว
tmp_dir = Path("/tmp/financial_engine")
tmp_dir.mkdir(exist_ok=True)
input_path = tmp_dir / uploaded_file.name
with open(input_path, "wb") as f:
    f.write(uploaded_file.getbuffer())

# โหลด Engine
engine_dir = Path(__file__).parent
sys.path.insert(0, str(engine_dir))

try:
    from financial_engine import FinancialEngine
    from export_9sheets import build_9sheets
except ImportError as e:
    st.error(f"❌ โหลด engine ไม่ได้: {e}\nตรวจสอบว่า financial_engine.py และ export_9sheets.py อยู่ในโฟลเดอร์เดียวกัน")
    st.stop()

with st.spinner("🔄 กำลังอ่านและวิเคราะห์ข้อมูล..."):
    try:
        engine = FinancialEngine(
            gemini_api_key  = gemini_key,
            target_year     = target_year,
            retained_begin  = retained_begin,
            admin_expense   = admin_expense,
            inventory_begin = inventory_begin,
            inventory_end   = inventory_end,
        )
        summary = engine.process(str(input_path), skip_gemini=not use_gemini)

        # Override ชื่อบริษัท
        summary.company.name   = company_name
        summary.company.tax_id = tax_id

        # แก้ total_sale = sale_vat + sale_exempt
        for m in summary.monthly:
            m.total_sale = m.sale_vat + m.sale_exempt

        # คำนวณใหม่
        summary.total_revenue = sum(m.total_sale for m in summary.monthly if m.total_sale > 0)
        summary.total_cogs    = inventory_begin + sum(m.purchase for m in summary.monthly) - inventory_end
        summary.gross_profit  = summary.total_revenue - summary.total_cogs
        summary.ebit          = summary.gross_profit - admin_expense

        def calc_cit(profit):
            if profit <= 0: return 0.
            tax = 0.
            for limit, rate, prev in [
                (300_000, 0., 0),
                (3_000_000, 0.15, 300_000),
                (float('inf'), 0.20, 3_000_000)
            ]:
                if profit > prev:
                    tax += (min(profit, limit) - prev) * rate
            return tax

        cit = calc_cit(max(summary.ebit, 0))
        summary.tax_expense  = -cit
        summary.net_profit   = summary.ebit - cit
        summary.retained_end = retained_begin + summary.net_profit

        st.session_state["summary"] = summary
        st.session_state["step_2"]  = True

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        st.exception(e)
        st.stop()

summary = st.session_state["summary"]
valid   = [m for m in summary.monthly if m.total_sale > 0 or m.vat_net != 0]


# ── Step 2: ตรวจสอบข้อมูล ────────────────────────────────────
st.markdown("---")
st.markdown("## 📋 ตรวจสอบข้อมูล")

# Metric cards
c1,c2,c3,c4,c5 = st.columns(5)
metrics = [
    (c1, "รายได้รวม",    summary.total_revenue, "฿", "normal"),
    (c2, "ต้นทุนขาย",    summary.total_cogs,    "฿", "normal"),
    (c3, "กำไรขั้นต้น",  summary.gross_profit,  "฿",
     "green" if summary.gross_profit > 0 else "red"),
    (c4, "กำไรสุทธิ",   summary.net_profit,    "฿",
     "green" if summary.net_profit > 0 else "red"),
    (c5, "VAT สุทธิ",   summary.total_vat_net, "฿", "normal"),
]
for col, label, val, unit, color in metrics:
    with col:
        pct = f" ({val/summary.total_revenue:.1%})" if summary.total_revenue > 0 and label != "รายได้รวม" and label != "VAT สุทธิ" else ""
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">{label}</div>
            <div class="value {color}">{val:,.0f}</div>
            <div class="label">{unit}{pct}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ตารางรายเดือน
st.markdown("### 📅 ข้อมูลรายเดือน")
import pandas as pd

month_rows = []
for m in summary.monthly:
    month_rows.append({
        "เดือน":       m.month_th,
        "ยอดขายรวม":   f"{m.total_sale:,.2f}" if m.total_sale > 0 else "-",
        "ยอดซื้อ":     f"{m.purchase:,.2f}"   if m.purchase > 0  else "-",
        "ภาษีขาย":    f"{m.vat_sale:,.2f}"    if m.vat_sale > 0  else "-",
        "ภาษีซื้อ":    f"{m.vat_buy:,.2f}"    if m.vat_buy > 0   else "-",
        "VAT ชำระ":    f"{m.vat_net:,.2f}",
        "Confidence":  "●" * int(m.confidence * 5) + "○" * (5 - int(m.confidence * 5)),
    })

df_months = pd.DataFrame(month_rows)
st.dataframe(df_months, use_container_width=True, hide_index=True)

# Anomalies
if summary.anomalies:
    st.markdown("### ⚠️ ความผิดปกติที่ตรวจพบ")
    for a in summary.anomalies:
        css_class = "anomaly-card high" if a['severity'] == 'high' else "anomaly-card"
        icon = "🔴" if a['severity'] == 'high' else "🟡"
        st.markdown(f"""
        <div class="{css_class}">
            {icon} <strong>[{a['severity'].upper()}] {a['month']}</strong>: {a['detail']}
        </div>""", unsafe_allow_html=True)
else:
    st.success("✅ ไม่พบความผิดปกติในข้อมูล")

# Gemini result
if use_gemini and summary.gemini_analysis:
    with st.expander("🤖 ผล Gemini Analysis"):
        verdict_color = "status-ok" if summary.gemini_verdict == "ready" else "status-warn"
        verdict_text  = "พร้อม Export" if summary.gemini_verdict == "ready" else "ควรตรวจสอบเพิ่ม"
        st.markdown(f'<p class="{verdict_color}">Verdict: {verdict_text}</p>',
                    unsafe_allow_html=True)
        st.text(summary.gemini_analysis)


# ── Feedback: ยืนยันก่อน Export ─────────────────────────────
st.markdown("---")
st.markdown("### ✅ ยืนยันข้อมูลก่อน Export")

col_confirm, col_note = st.columns([1,2])
with col_confirm:
    confirmed = st.checkbox(
        "ตรวจสอบตัวเลขแล้ว พร้อม Export",
        help="กดเพื่อยืนยันว่าตัวเลขถูกต้อง"
    )
with col_note:
    note = st.text_input(
        "หมายเหตุ (ถ้ามี)",
        placeholder="เช่น ตรวจสอบแล้ว ถูกต้อง / ยอดเดือน ม.ค. ต้องตรวจสอบใหม่"
    )

# บันทึก feedback ลง training data
if confirmed and note:
    feedback_dir = Path("training_data")
    feedback_dir.mkdir(exist_ok=True)
    feedback = {
        "timestamp":    datetime.now().isoformat(),
        "company":      company_name,
        "year":         target_year,
        "file":         uploaded_file.name,
        "is_correct":   True,
        "note":         note,
        "metrics": {
            "total_revenue": summary.total_revenue,
            "net_profit":    summary.net_profit,
            "vat_net":       summary.total_vat_net,
        }
    }
    with open(feedback_dir / "feedback_log.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps(feedback, ensure_ascii=False) + "\n")


# ── Step 3: Export ───────────────────────────────────────────
st.markdown("---")
st.markdown("## 📥 Export Excel งบการเงิน 9 Sheets")

# ข้อมูลปีก่อน
with st.expander("📌 ข้อมูลปีก่อน (2567) — ใส่เพื่อเปรียบเทียบใน PL"):
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        prev_revenue    = st.number_input("รายได้จากการขาย 2567",  value=8715687.39, format="%.2f")
        prev_other      = st.number_input("รายได้อื่น 2567",         value=20000.00,   format="%.2f")
        prev_cogs       = st.number_input("ต้นทุนขาย 2567",          value=7807814.99, format="%.2f")
        prev_admin      = st.number_input("ค่าใช้จ่ายบริหาร 2567",  value=565280.31,  format="%.2f")
    with col_p2:
        prev_ebit       = st.number_input("กำไรก่อนภาษี 2567",      value=362592.09,  format="%.2f")
        prev_tax        = st.number_input("ภาษีเงินได้ 2567",        value=-9631.86,   format="%.2f")
        prev_net        = st.number_input("กำไรสุทธิ 2567",          value=352960.23,  format="%.2f")
        prev_ret_begin  = st.number_input("กำไรสะสมต้นปี 2567",     value=1925166.83, format="%.2f")
        prev_ret_end    = st.number_input("กำไรสะสมปลายปี 2567",    value=2278127.06, format="%.2f")

prev_year_data = {
    'revenue':       prev_revenue,
    'other_income':  prev_other,
    'total_revenue': prev_revenue + prev_other,
    'cogs':          prev_cogs,
    'admin':         prev_admin,
    'total_exp':     prev_cogs + prev_admin,
    'ebit':          prev_ebit,
    'tax':           prev_tax,
    'net_profit':    prev_net,
    'retained_begin':prev_ret_begin,
    'retained_end':  prev_ret_end,
}

col_export, col_info = st.columns([1,2])
with col_export:
    export_btn = st.button(
        "📊 สร้าง Excel 9 Sheets",
        disabled=not confirmed,
        use_container_width=True,
    )

with col_info:
    if not confirmed:
        st.warning("⚠️ กรุณาติ๊กยืนยันข้อมูลก่อน")
    else:
        st.success("✅ พร้อม Export")

if export_btn and confirmed:
    output_path = tmp_dir / f"งบการเงิน_{company_name.replace(' ','_')}_25{target_year}.xlsx"

    with st.spinner("⏳ กำลังสร้าง Excel..."):
        try:
            build_9sheets(
                summary,
                str(output_path),
                manager_name    = manager_name,
                prev_year_label = f"25{int(target_year)-1}",
                prev_year_data  = prev_year_data,
            )
            st.session_state["output_path"] = str(output_path)
            st.session_state["step_3"] = True
            st.balloons()
            st.success(f"✅ สร้าง Excel สำเร็จ!")
        except Exception as e:
            st.error(f"❌ Export ไม่สำเร็จ: {e}")
            st.exception(e)

# Download button
if st.session_state.get("output_path") and Path(st.session_state["output_path"]).exists():
    output_path = st.session_state["output_path"]
    fname = Path(output_path).name

    with open(output_path, "rb") as f:
        file_bytes = f.read()

    st.download_button(
        label     = f"⬇️  Download {fname}",
        data      = file_bytes,
        file_name = fname,
        mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width = True,
    )

    st.markdown("### 📋 สรุป Sheets ที่ได้")
    sheets_info = [
        ("งบทดลอง",        "auto บางส่วน + template"),
        ("งบดุล",          "template (กำไรสะสม auto)"),
        ("PL",             "✅ auto จาก ภพ.30"),
        ("หมายเหตุฯ",      "auto + template"),
        (f"ภพ.30.{target_year}", "✅ auto จาก ภพ.30"),
        ("ภาษีที่ต้องชำระ","✅ auto + template"),
        ("เงินเดือน",       "⚠️  กรอกเอง"),
        ("สินค้าคงเหลือ",   "⚠️  กรอกเอง"),
        ("ต้นทุนข้ามรอบ",   "⚠️  กรอกเอง"),
    ]
    for i, (sheet, status) in enumerate(sheets_info, 1):
        icon = "✅" if "auto" in status and "⚠️" not in status else "⚠️" if "⚠️" in status else "📄"
        st.markdown(f"**{i}. {sheet}** — {status}")


# ── Footer ───────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<p style="text-align:center;color:#9CA3AF;font-size:0.85rem;">'
    'Financial Engine v1.0  |  Powered by Gemini AI  |  '
    f'© {datetime.now().year}</p>',
    unsafe_allow_html=True
)
