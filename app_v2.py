"""
app_v2.py — Financial Engine Web App
Phase 1: Home + Company Profiles + จำงบแต่ละปี
"""
import streamlit as st
import sys, os, json, time
from pathlib import Path
from datetime import datetime

st.set_page_config(
    page_title="Financial Engine",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');
html,body,[class*="css"]{font-family:'Sarabun',sans-serif!important}
.hero{background:linear-gradient(135deg,#1E3A5F,#1A56DB);padding:2rem 2.5rem;
      border-radius:16px;margin-bottom:1.5rem;box-shadow:0 8px 32px rgba(26,86,219,.25)}
.hero h1{color:#fff;font-size:1.9rem;font-weight:700;margin:0}
.hero p{color:#BFD9F5;margin:.25rem 0 0;font-size:1rem}
.co-card{background:#F8FAFF;border:1.5px solid #DBEAFE;border-radius:14px;
         padding:1.2rem 1.5rem;cursor:pointer;transition:all .2s;margin-bottom:.75rem}
.co-card:hover{border-color:#1A56DB;box-shadow:0 4px 16px rgba(26,86,219,.12);transform:translateY(-2px)}
.co-card h3{color:#1E3A5F;font-size:1rem;margin:0 0 .25rem}
.co-card p{color:#6B7280;font-size:.85rem;margin:0}
.co-card .badge{display:inline-block;background:#DBEAFE;color:#1A56DB;
                border-radius:20px;padding:2px 10px;font-size:.75rem;font-weight:600}
.year-chip{display:inline-block;background:#ECFDF5;color:#065F46;border-radius:10px;
           padding:3px 10px;font-size:.8rem;font-weight:600;margin:2px}
.year-chip.selected{background:#065F46;color:#fff}
.metric-box{background:#F8FAFF;border:1px solid #DBEAFE;border-radius:10px;
            padding:1rem;text-align:center}
.metric-box .lbl{color:#6B7280;font-size:.8rem}
.metric-box .val{color:#1E3A5F;font-size:1.3rem;font-weight:700}
.metric-box .val.g{color:#065F46}.metric-box .val.r{color:#DC2626}
.warn-box{background:#FEF3C7;border-left:4px solid #F59E0B;border-radius:0 8px 8px 0;
          padding:.75rem 1rem;margin:.4rem 0}
.warn-box.high{background:#FEF2F2;border-color:#DC2626}
div.stButton>button{background:#1A56DB;color:#fff;border:none;border-radius:8px;
    padding:.55rem 1.8rem;font-weight:600;font-family:'Sarabun',sans-serif;
    font-size:1rem;transition:all .2s}
div.stButton>button:hover{background:#1E3A5F;transform:translateY(-1px)}
div.stButton>button:disabled{background:#D1D5DB;cursor:not-allowed;transform:none}
</style>
""", unsafe_allow_html=True)

# ── Storage: จำงบแต่ละบริษัทแต่ละปี ─────────────────────────
DATA_DIR = Path("company_data")
DATA_DIR.mkdir(exist_ok=True)

def load_companies() -> dict:
    f = DATA_DIR / "companies.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}

def save_companies(data: dict):
    (DATA_DIR / "companies.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def save_yearly_report(tax_id: str, year: str, report: dict):
    f = DATA_DIR / f"{tax_id}_{year}.json"
    f.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

def load_yearly_report(tax_id: str, year: str) -> dict:
    f = DATA_DIR / f"{tax_id}_{year}.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}

def get_company_years(tax_id: str) -> list:
    return sorted([
        f.stem.split("_")[1]
        for f in DATA_DIR.glob(f"{tax_id}_*.json")
        if f.stem != "companies"
    ], reverse=True)

companies = load_companies()

# ── Session state ────────────────────────────────────────────
if "page" not in st.session_state:
    st.session_state.page = "home"
if "selected_company" not in st.session_state:
    st.session_state.selected_company = None
if "summary" not in st.session_state:
    st.session_state.summary = None
if "output_path" not in st.session_state:
    st.session_state.output_path = None


# ════════════════════════════════════════════════════════════
# PAGE: HOME
# ════════════════════════════════════════════════════════════
def page_home():
    st.markdown("""
    <div class="hero">
        <h1>📊 Financial Engine</h1>
        <p>ระบบปิดงบการเงินอัตโนมัติจาก ภ.พ.30 — เลือกบริษัทหรือเพิ่มใหม่</p>
    </div>""", unsafe_allow_html=True)

    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.markdown("### 🏢 บริษัทที่มีในระบบ")
        if not companies:
            st.info("ยังไม่มีบริษัทในระบบ — กด **+ เพิ่มบริษัทใหม่** ด้านขวา")
        else:
            for tax_id, co in companies.items():
                years = get_company_years(tax_id)
                year_chips = " ".join([f'<span class="year-chip">{y}</span>' for y in years[:5]])
                if st.button(f"**{co['name']}**\n\n{co['tax_id']}", key=f"co_{tax_id}",
                             use_container_width=True):
                    st.session_state.selected_company = tax_id
                    st.session_state.page = "company"
                    st.rerun()

                # แสดงปีที่มีข้อมูล
                if years:
                    st.markdown(
                        f'<div style="margin:-10px 0 8px;padding-left:4px">'
                        f'ปีที่มีข้อมูล: {year_chips}</div>',
                        unsafe_allow_html=True)

    with col_right:
        st.markdown("### ➕ เพิ่มบริษัทใหม่")
        with st.form("add_company"):
            new_name    = st.text_input("ชื่อบริษัท", placeholder="หจก. / บจก. ...")
            new_tax_id  = st.text_input("เลขผู้เสียภาษี", placeholder="0000000000000")
            new_manager = st.text_input("ชื่อผู้จัดการ/หุ้นส่วน")
            new_biz     = st.selectbox("ประเภทธุรกิจ",
                ["ทั่วไป","ค้าปลีก","ค้าส่ง","บริการ","รถยนต์","ก่อสร้าง","อาหาร"])
            submitted = st.form_submit_button("✅ เพิ่มบริษัท", use_container_width=True)
            if submitted:
                if not new_name or not new_tax_id:
                    st.error("กรุณากรอกชื่อและเลขผู้เสียภาษี")
                elif new_tax_id in companies:
                    st.warning("บริษัทนี้มีในระบบแล้ว")
                else:
                    companies[new_tax_id] = {
                        "name": new_name, "tax_id": new_tax_id,
                        "manager": new_manager, "business_type": new_biz,
                        "created_at": datetime.now().isoformat(),
                    }
                    save_companies(companies)
                    st.success(f"✅ เพิ่ม {new_name} แล้ว")
                    st.rerun()


# ════════════════════════════════════════════════════════════
# PAGE: COMPANY (เลือกปี + ดูประวัติ)
# ════════════════════════════════════════════════════════════
def page_company():
    tax_id = st.session_state.selected_company
    if not tax_id or tax_id not in companies:
        st.session_state.page = "home"; st.rerun()

    co    = companies[tax_id]
    years = get_company_years(tax_id)

    # Breadcrumb
    col_back, col_title = st.columns([1, 5])
    with col_back:
        if st.button("← กลับ"):
            st.session_state.page = "home"; st.rerun()
    with col_title:
        st.markdown(f"### 🏢 {co['name']}")
        st.caption(f"เลขผู้เสียภาษี: {co['tax_id']}  |  {co.get('business_type','')}")

    st.markdown("---")
    col_new, col_hist = st.columns([1, 1])

    # ── ปิดงบปีใหม่ ──────────────────────────────────────────
    with col_new:
        st.markdown("#### 📂 ปิดงบปีใหม่")
        target_year = st.selectbox("ปี พ.ศ. (2 หลัก)", ["68","69","67"], key="new_year")

        # ดึงข้อมูลปีก่อนอัตโนมัติ
        prev_year = str(int(target_year) - 1)
        prev_report = load_yearly_report(tax_id, prev_year)
        if prev_report:
            st.success(f"✅ พบข้อมูลปี {prev_year} — จะใช้เปรียบเทียบอัตโนมัติ")

        if st.button("🚀 เริ่มปิดงบ", use_container_width=True, key="start_new"):
            st.session_state.page = "process"
            st.session_state.process_year = target_year
            st.rerun()

    # ── ประวัติงบที่ปิดไว้ ─────────────────────────────────────
    with col_hist:
        st.markdown("#### 📋 ประวัติงบที่ปิดไว้")
        if not years:
            st.info("ยังไม่มีประวัติ")
        else:
            for yr in years:
                rpt = load_yearly_report(tax_id, yr)
                if rpt:
                    rev = rpt.get("total_revenue", 0)
                    net = rpt.get("net_profit", 0)
                    col_y, col_r, col_n, col_btn = st.columns([1,2,2,1])
                    col_y.markdown(f"**25{yr}**")
                    col_r.markdown(f"<small>รายได้</small><br>{rev:,.0f}", unsafe_allow_html=True)
                    col_n.markdown(
                        f"<small>กำไรสุทธิ</small><br>"
                        f"<span style='color:{'#065F46' if net>=0 else '#DC2626'}'>{net:,.0f}</span>",
                        unsafe_allow_html=True)
                    if col_btn.button("ดู", key=f"view_{yr}"):
                        st.session_state.page = "view_report"
                        st.session_state.view_year = yr
                        st.rerun()

    # ── เปรียบเทียบหลายปี ──────────────────────────────────────
    if len(years) >= 2:
        st.markdown("---")
        st.markdown("#### 📈 เปรียบเทียบหลายปี")
        import pandas as pd
        rows = []
        for yr in years[:5]:
            r = load_yearly_report(tax_id, yr)
            if r:
                rows.append({
                    "ปี": f"25{yr}",
                    "รายได้": f"{r.get('total_revenue',0):,.0f}",
                    "ต้นทุนขาย": f"{r.get('total_cogs',0):,.0f}",
                    "กำไรขั้นต้น": f"{r.get('gross_profit',0):,.0f}",
                    "กำไรสุทธิ": f"{r.get('net_profit',0):,.0f}",
                    "VAT สุทธิ": f"{r.get('total_vat_net',0):,.0f}",
                })
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════════════
# PAGE: PROCESS (อ่าน ภพ.30 + Export)
# ════════════════════════════════════════════════════════════
def page_process():
    tax_id = st.session_state.selected_company
    co     = companies.get(tax_id, {})
    yr     = st.session_state.get("process_year", "68")

    col_back, col_title = st.columns([1, 5])
    with col_back:
        if st.button("← กลับ"):
            st.session_state.page = "company"; st.rerun()
    with col_title:
        st.markdown(f"### 📊 ปิดงบ 25{yr} — {co.get('name','')}")

    st.markdown("---")

    # Sidebar params
    with st.sidebar:
        st.markdown(f"### ⚙️ ตั้งค่า ปี 25{yr}")
        retained_begin  = st.number_input("กำไรสะสมต้นปี", value=0.0, format="%.2f")
        admin_expense   = st.number_input("ค่าใช้จ่ายบริหาร", value=0.0, format="%.2f")
        inventory_begin = st.number_input("สินค้าต้นงวด", value=0.0, format="%.2f")
        inventory_end   = st.number_input("สินค้าปลายงวด", value=0.0, format="%.2f")
        use_gemini      = st.toggle("Gemini วิเคราะห์", value=False)
        gemini_key      = os.environ.get("GEMINI_API_KEY", "dummy")
        if use_gemini:
            gemini_key  = st.text_input("Gemini API Key", type="password")

        # ดึงข้อมูลปีก่อนอัตโนมัติ
        prev_yr    = str(int(yr) - 1)
        prev_rpt   = load_yearly_report(tax_id, prev_yr)
        if prev_rpt:
            st.success(f"✅ ใช้ข้อมูลปี {prev_yr} เปรียบเทียบ")
            retained_begin = prev_rpt.get("retained_end", 0.0)
            st.info(f"กำไรสะสมต้นปี: {retained_begin:,.2f}")

    # Steps
    step_cols = st.columns(3)
    for i,(icon,label) in enumerate([("1️⃣","อัปโหลดไฟล์"),
                                      ("2️⃣","ตรวจสอบข้อมูล"),
                                      ("3️⃣","Export Excel")]):
        active = st.session_state.get(f"step_{i+1}", False)
        with step_cols[i]:
            st.markdown(
                f'<div style="background:{"#1E3A5F" if active else "#F3F4F6"};'
                f'color:{"#fff" if active else "#9CA3AF"};border-radius:10px;'
                f'padding:.7rem;text-align:center;font-weight:600">{icon} {label}</div>',
                unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Upload
    uploaded = st.file_uploader(
        "📂 อัปโหลดไฟล์ ภ.พ.30 (.xlsx หรือ .xls)",
        type=["xlsx","xls"])

    if not uploaded:
        st.info("👆 อัปโหลดไฟล์ ภ.พ.30 เพื่อเริ่มต้น")
        return

    st.session_state["step_1"] = True

    # Save temp file
    tmp = Path("/tmp/financial_engine"); tmp.mkdir(exist_ok=True)
    input_path = tmp / uploaded.name
    input_path.write_bytes(uploaded.getbuffer())

    # Load engine
    engine_dir = Path(__file__).parent
    sys.path.insert(0, str(engine_dir))
    try:
        from financial_engine import FinancialEngine
        from export_9sheets import build_9sheets
    except ImportError as e:
        st.error(f"❌ โหลด engine ไม่ได้: {e}")
        return

    with st.spinner("🔄 กำลังอ่านและวิเคราะห์..."):
        try:
            engine = FinancialEngine(
                gemini_api_key  = gemini_key,
                target_year     = yr,
                retained_begin  = retained_begin,
                admin_expense   = admin_expense,
                inventory_begin = inventory_begin,
                inventory_end   = inventory_end,
            )
            summary = engine.process(str(input_path), skip_gemini=not use_gemini)
            summary.company.name   = co["name"]
            summary.company.tax_id = co["tax_id"]

            for m in summary.monthly:
                m.total_sale = m.sale_vat + m.sale_exempt

            summary.total_revenue = sum(m.total_sale for m in summary.monthly if m.total_sale > 0)
            summary.total_cogs    = inventory_begin + sum(m.purchase for m in summary.monthly) - inventory_end
            summary.gross_profit  = summary.total_revenue - summary.total_cogs
            summary.ebit          = summary.gross_profit - admin_expense

            def cit(p):
                if p<=0: return 0.
                t=0.
                for lim,r,pv in [(300_000,0.,0),(3_000_000,.15,300_000),(float('inf'),.20,3_000_000)]:
                    if p>pv: t+=(min(p,lim)-pv)*r
                return t
            c = cit(max(summary.ebit,0))
            summary.tax_expense  = -c
            summary.net_profit   = summary.ebit - c
            summary.retained_end = retained_begin + summary.net_profit

            st.session_state.summary      = summary
            st.session_state["step_2"]    = True
        except Exception as e:
            st.error(f"❌ {e}"); st.exception(e); return

    summary = st.session_state.summary
    valid   = [m for m in summary.monthly if m.total_sale > 0 or m.vat_net != 0]

    # Metrics
    st.markdown("---")
    st.markdown("## 📋 ตรวจสอบข้อมูล")
    mc = st.columns(5)
    for col, label, val, cls in zip(mc, [
        "รายได้รวม","ต้นทุนขาย","กำไรขั้นต้น","กำไรสุทธิ","VAT สุทธิ"
    ],[
        summary.total_revenue, summary.total_cogs, summary.gross_profit,
        summary.net_profit, summary.total_vat_net
    ],[
        "","","g" if summary.gross_profit>0 else "r",
        "g" if summary.net_profit>0 else "r",""
    ]):
        pct = (f" ({val/summary.total_revenue:.1%})"
               if summary.total_revenue > 0 and label not in ["รายได้รวม","VAT สุทธิ"]
               else "")
        with col:
            st.markdown(
                f'<div class="metric-box"><div class="lbl">{label}</div>'
                f'<div class="val {cls}">{val:,.0f}</div>'
                f'<div class="lbl">฿{pct}</div></div>',
                unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Monthly table
    import pandas as pd
    rows = [{
        "เดือน":      m.month_th,
        "ยอดขายรวม":  f"{m.total_sale:,.2f}" if m.total_sale > 0 else "-",
        "ยอดซื้อ":    f"{m.purchase:,.2f}" if m.purchase > 0 else "-",
        "ภาษีขาย":   f"{m.vat_sale:,.2f}" if m.vat_sale > 0 else "-",
        "VAT ชำระ":   f"{m.vat_net:,.2f}",
        "Conf": "●"*int(m.confidence*5)+"○"*(5-int(m.confidence*5)),
    } for m in summary.monthly]
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    if summary.anomalies:
        st.markdown("### ⚠️ ความผิดปกติ")
        for a in summary.anomalies:
            css = "warn-box high" if a["severity"]=="high" else "warn-box"
            st.markdown(
                f'<div class="{css}"><b>[{a["severity"].upper()}] {a["month"]}</b>: {a["detail"]}</div>',
                unsafe_allow_html=True)

    # Confirm + export
    st.markdown("---")
    confirmed = st.checkbox("✅ ตรวจสอบตัวเลขแล้ว พร้อม Export")
    note      = st.text_input("หมายเหตุ", placeholder="เช่น ตรวจสอบแล้ว ถูกต้อง")

    # Prev year data
    with st.expander("📌 ข้อมูลปีก่อน (เปรียบเทียบใน PL)"):
        if prev_rpt:
            st.success(f"✅ โหลดข้อมูลปี {prev_yr} อัตโนมัติ")
            prev_data = {
                "revenue":        prev_rpt.get("total_revenue",0),
                "other_income":   prev_rpt.get("other_income",0),
                "total_revenue":  prev_rpt.get("total_revenue",0),
                "cogs":           prev_rpt.get("total_cogs",0),
                "admin":          prev_rpt.get("admin_expense",0),
                "total_exp":      prev_rpt.get("total_cogs",0)+prev_rpt.get("admin_expense",0),
                "ebit":           prev_rpt.get("ebit",0),
                "tax":            prev_rpt.get("tax_expense",0),
                "net_profit":     prev_rpt.get("net_profit",0),
                "retained_begin": prev_rpt.get("retained_begin",0),
                "retained_end":   prev_rpt.get("retained_end",0),
            }
        else:
            st.info(f"ไม่พบข้อมูลปี {prev_yr} — กรอกเองด้านล่าง")
            c1,c2 = st.columns(2)
            with c1:
                p_rev  = st.number_input("รายได้",     value=0.0, format="%.2f")
                p_cogs = st.number_input("ต้นทุนขาย",  value=0.0, format="%.2f")
                p_adm  = st.number_input("ค่าบริหาร",  value=0.0, format="%.2f")
            with c2:
                p_ebit = st.number_input("กำไรก่อนภาษี",value=0.0, format="%.2f")
                p_tax  = st.number_input("ภาษีเงินได้", value=0.0, format="%.2f")
                p_net  = st.number_input("กำไรสุทธิ",   value=0.0, format="%.2f")
                p_rbe  = st.number_input("กำไรสะสมต้นปี",value=0.0,format="%.2f")
                p_ren  = st.number_input("กำไรสะสมปลายปี",value=0.0,format="%.2f")
            prev_data = {
                "revenue":p_rev,"other_income":0.,"total_revenue":p_rev,
                "cogs":p_cogs,"admin":p_adm,"total_exp":p_cogs+p_adm,
                "ebit":p_ebit,"tax":p_tax,"net_profit":p_net,
                "retained_begin":p_rbe,"retained_end":p_ren,
            }

    if st.button("📊 สร้าง Excel 9 Sheets", disabled=not confirmed,
                 use_container_width=True):
        output_path = tmp / f"งบการเงิน_{co['name'].replace(' ','_')}_25{yr}.xlsx"
        with st.spinner("⏳ กำลังสร้าง..."):
            try:
                build_9sheets(summary, str(output_path),
                              manager_name    = co.get("manager",""),
                              prev_year_label = f"25{prev_yr}",
                              prev_year_data  = prev_data)
                st.session_state.output_path = str(output_path)
                st.session_state["step_3"]   = True

                # บันทึกงบลง storage
                save_yearly_report(tax_id, yr, {
                    "company":       co["name"],
                    "tax_id":        co["tax_id"],
                    "year":          yr,
                    "closed_at":     datetime.now().isoformat(),
                    "note":          note,
                    "total_revenue": round(summary.total_revenue,2),
                    "total_cogs":    round(summary.total_cogs,2),
                    "gross_profit":  round(summary.gross_profit,2),
                    "admin_expense": admin_expense,
                    "ebit":          round(summary.ebit,2),
                    "tax_expense":   round(summary.tax_expense,2),
                    "net_profit":    round(summary.net_profit,2),
                    "retained_begin":retained_begin,
                    "retained_end":  round(summary.retained_end,2),
                    "total_vat_net": round(summary.total_vat_net,2),
                    "other_income":  0.,
                    "monthly": [{
                        "month": m.month_th,
                        "total_sale": round(m.total_sale,2),
                        "purchase":   round(m.purchase,2),
                        "vat_net":    round(m.vat_net,2),
                    } for m in summary.monthly],
                })
                st.balloons()
                st.success("✅ ปิดงบสำเร็จ! บันทึกในระบบแล้ว")
            except Exception as e:
                st.error(f"❌ {e}"); st.exception(e)

    if st.session_state.output_path and Path(st.session_state.output_path).exists():
        with open(st.session_state.output_path,"rb") as f:
            st.download_button(
                "⬇️ Download Excel งบการเงิน 9 Sheets",
                data      = f.read(),
                file_name = Path(st.session_state.output_path).name,
                mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width = True,
            )


# ════════════════════════════════════════════════════════════
# PAGE: VIEW REPORT
# ════════════════════════════════════════════════════════════
def page_view_report():
    tax_id = st.session_state.selected_company
    yr     = st.session_state.get("view_year","")
    rpt    = load_yearly_report(tax_id, yr)
    co     = companies.get(tax_id, {})

    col_back, col_title = st.columns([1,5])
    with col_back:
        if st.button("← กลับ"):
            st.session_state.page = "company"; st.rerun()
    with col_title:
        st.markdown(f"### 📋 งบการเงินปี 25{yr} — {co.get('name','')}")
        st.caption(f"ปิดงบเมื่อ: {rpt.get('closed_at','')[:10]}  |  {rpt.get('note','')}")

    if not rpt:
        st.warning("ไม่พบข้อมูล")
        return

    mc = st.columns(5)
    for col, label, val, cls in zip(mc,
        ["รายได้รวม","ต้นทุนขาย","กำไรขั้นต้น","กำไรสุทธิ","VAT สุทธิ"],
        [rpt.get("total_revenue",0), rpt.get("total_cogs",0),
         rpt.get("gross_profit",0),  rpt.get("net_profit",0),
         rpt.get("total_vat_net",0)],
        ["","","g","g" if rpt.get("net_profit",0)>=0 else "r",""],
    ):
        with col:
            st.markdown(
                f'<div class="metric-box"><div class="lbl">{label}</div>'
                f'<div class="val {cls}">{val:,.0f}</div>'
                f'<div class="lbl">฿</div></div>',
                unsafe_allow_html=True)

    if rpt.get("monthly"):
        import pandas as pd
        st.markdown("### รายเดือน")
        rows = [{"เดือน":m["month"],
                 "ยอดขาย":f"{m['total_sale']:,.0f}",
                 "ยอดซื้อ":f"{m['purchase']:,.0f}",
                 "VAT":f"{m['vat_net']:,.2f}"}
                for m in rpt["monthly"]]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════════════
# ROUTER
# ════════════════════════════════════════════════════════════
page = st.session_state.page
if page == "home":
    page_home()
elif page == "company":
    page_company()
elif page == "process":
    page_process()
elif page == "view_report":
    page_view_report()
