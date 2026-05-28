"""
financial_engine.py
═══════════════════════════════════════════════════════════════
Core Engine สำหรับวิเคราะห์และปิดงบการเงิน จากไฟล์ ภพ.30
รองรับโครงสร้างหลากหลายของบริษัทต่างๆ

โครงสร้างหลัก:
  1. VAT30Reader      — อ่านและ parse ภพ.30 ทุกรูปแบบ
  2. AccountClassifier — คัดแยกและจัดหมวดบัญชีอัตโนมัติ
  3. FinancialCalc     — คำนวณงบการเงิน (BS / PL / VAT)
  4. GeminiAnalyzer   — ส่งให้ Gemini วิเคราะห์และตรวจสอบ
  5. FinancialEngine  — orchestrator รวมทุก module

วิธีใช้:
  engine = FinancialEngine(gemini_api_key="...", target_year="68")
  result = engine.process("ภพ.30_บริษัทXYZ.xlsx")
  result.export("output.xlsx")
═══════════════════════════════════════════════════════════════
"""

import re
import time
import json
import logging
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional

import pandas as pd
import numpy as np
import google.generativeai as genai
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════
# DATA STRUCTURES
# ══════════════════════════════════════════════════════════════

@dataclass
class CompanyInfo:
    """ข้อมูลบริษัท"""
    name: str = ""
    tax_id: str = ""
    address: str = ""
    business_type: str = ""
    fiscal_year: str = ""

@dataclass
class MonthlyVAT:
    """ข้อมูลภาษีรายเดือน"""
    month_num: int = 0
    month_th: str = ""
    month_full: str = ""
    # ยอดขาย
    sale_vat: float = 0.0        # ขายมี VAT
    sale_exempt: float = 0.0     # ขายยกเว้น VAT
    sale_zero: float = 0.0       # ขาย 0%
    total_sale: float = 0.0      # รวมยอดขายทั้งหมด
    # ยอดซื้อ
    purchase: float = 0.0        # ยอดซื้อ
    purchase_exempt: float = 0.0 # ซื้อยกเว้น VAT
    total_purchase: float = 0.0  # รวมยอดซื้อทั้งหมด
    # ภาษี
    vat_sale: float = 0.0        # ภาษีขาย
    vat_buy: float = 0.0         # ภาษีซื้อ
    vat_net: float = 0.0         # ภาษีที่ต้องชำระ (บวก=ชำระ, ลบ=ขอคืน)
    vat_cumul: float = 0.0       # ภาษีสะสม
    vat_penalty: float = 0.0     # เบี้ยปรับ
    # กำไร
    gross_profit: float = 0.0
    gross_profit_rate: float = 0.0
    # Meta
    source_sheet: str = ""
    sheet_type: str = ""         # 'single'|'pair'|'summary'
    confidence: float = 1.0      # ความมั่นใจใน extraction (0-1)
    notes: str = ""

@dataclass
class FinancialSummary:
    """สรุปงบการเงินทั้งปี"""
    company: CompanyInfo = field(default_factory=CompanyInfo)
    year: str = ""
    monthly: list = field(default_factory=list)  # list[MonthlyVAT]
    # งบกำไรขาดทุน
    total_revenue: float = 0.0
    total_cogs: float = 0.0
    gross_profit: float = 0.0
    admin_expense: float = 0.0
    ebit: float = 0.0
    tax_expense: float = 0.0
    net_profit: float = 0.0
    retained_begin: float = 0.0
    retained_end: float = 0.0
    # ภาษี
    total_vat_sale: float = 0.0
    total_vat_buy: float = 0.0
    total_vat_net: float = 0.0
    # วิเคราะห์
    gemini_analysis: str = ""
    gemini_verdict: str = ""     # 'ready'|'review'|'error'
    anomalies: list = field(default_factory=list)


# ══════════════════════════════════════════════════════════════
# 1. VAT30Reader — อ่าน ภพ.30 ทุกรูปแบบ
# ══════════════════════════════════════════════════════════════

class VAT30Reader:
    """
    อ่านและ parse ไฟล์ ภพ.30 รองรับหลายโครงสร้าง:
      - โครงสร้าง A: แต่ละเดือน = 1 sheet (1.68, 2.68, ...)
      - โครงสร้าง B: แยก สนง/สาขา (01.68 สนง, 01.68 สาขา)
      - โครงสร้าง C: มี sheet สรุปรวม (01.68 สรุป)
      - โครงสร้าง D: ผสมหลายรูปแบบในไฟล์เดียว
    """

    MONTHS_TH   = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                   'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
    MONTHS_FULL = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                   'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม']

    # Keywords สำหรับ detect ประเภท sheet
    SUMMARY_KW = ['สรุป', 'summary', 'รวม']
    BRANCH_KW  = ['สาขา', 'branch', 'วาริน', 'สีคิ้ว', 'สาขา']
    OFFICE_KW  = ['สนง', 'สำนักงาน', 'office', 'ใหญ่']
    SKIP_KW    = ['(2)', 'ตาม', 'ปรับ', 'แก้', 'old', 'backup']

    # Keywords สำหรับหา summary rows
    TOTAL_SALE_KW   = ['รวมรายได้', 'รวมยอดขาย', 'total sale', 'ยอดขายรวม']
    TOTAL_BUY_KW    = ['รวมค่าใช้จ่าย', 'รวมยอดซื้อ', 'total purchase', 'ยอดซื้อรวม']
    EXEMPT_KW       = ['ยกเว้น', 'exempt', 'ยอดขายที่ได้รับยกเว้น']
    VAT_PAY_KW      = ['รวมชำระ ภ.พ.30', 'ภาษีที่ต้องชำระ', 'ต้องชำระ (ชำระเกินยกไป)',
                       'รวมชำระทั้งหมด', 'vat payable', 'ภาษีที่ต้องชำระเดือนนี้']
    PROFIT_RATE_KW  = ['อัตรากำไรขั้นต้น', 'gross profit rate', 'อัตรากำไร']
    COMPANY_KW      = ['ชื่อผู้ประกอบการ', 'ชื่อบริษัท', 'company name', 'ห้าง', 'บริษัท', 'หจก']
    TAXID_KW        = ['เลขประจำตัวผู้เสียภาษี', 'เลขทะเบียน', 'tax id', 'เลขที่ผู้เสียภาษี']

    def __init__(self, target_year: str):
        self.target_year = target_year  # เช่น "68"
        self.filepath = None
        self.xl = None
        self.sheets = []
        self._sheet_map = {}  # month_num -> sheet info

    def load(self, filepath: str) -> 'VAT30Reader':
        """โหลดไฟล์และ detect โครงสร้าง"""
        self.filepath = filepath
        ext = Path(filepath).suffix.lower()
        engine = 'xlrd' if ext == '.xls' else None
        self.xl = pd.ExcelFile(filepath, engine=engine)
        self.sheets = self.xl.sheet_names
        self._sheet_map = self._map_sheets()
        log.info(f"โหลดไฟล์: {filepath} | พบ {len(self.sheets)} sheets")
        log.info(f"Sheets ปี {self.target_year}: {list(self._sheet_map.keys())}")
        return self

    def _read(self, sheet_name: str) -> pd.DataFrame:
        """อ่าน sheet เป็น DataFrame"""
        ext = Path(self.filepath).suffix.lower()
        engine = 'xlrd' if ext == '.xls' else None
        return pd.read_excel(self.filepath, sheet_name=sheet_name,
                             header=None, engine=engine)

    def _map_sheets(self) -> dict:
        """
        จับคู่ sheet กับเดือน รองรับ pattern หลายแบบ:
          1.68 / 01.68 / 1.2568 / ม.ค.68 / Jan68
        """
        result = {}
        year = self.target_year
        year_long = f"25{year}"  # เช่น 2568

        for m in range(1, 13):
            m_pads = [str(m), f"{m:02d}"]
            candidates = []

            for s in self.sheets:
                s_lower = s.lower().strip()
                # skip backup/duplicate sheets
                if any(kw in s_lower for kw in [k.lower() for k in self.SKIP_KW]):
                    continue
                # match pattern
                matched = False
                for mp in m_pads:
                    if (f"{mp}.{year}" in s or
                        f"{mp}.{year_long}" in s or
                        f"{mp}/{year}" in s):
                        matched = True
                        break
                if matched:
                    candidates.append(s)

            if not candidates:
                continue

            # จัดหมวด candidates
            summary = [s for s in candidates if any(kw in s for kw in self.SUMMARY_KW)]
            branch  = [s for s in candidates if any(kw in s for kw in self.BRANCH_KW)
                       and s not in summary]
            office  = [s for s in candidates if any(kw in s for kw in self.OFFICE_KW)
                       and s not in summary and s not in branch]
            single  = [s for s in candidates if s not in summary + branch + office]

            if summary:
                result[m] = {'type': 'summary', 'sheets': summary[:1]}
            elif branch or office:
                result[m] = {'type': 'pair',
                             'sheets': list(set(branch + office + single))}
            elif single:
                result[m] = {'type': 'single', 'sheets': single[:1]}

        return result

    def get_company_info(self) -> CompanyInfo:
        """ดึงข้อมูลบริษัทจาก sheet แรกที่เจอ"""
        info = CompanyInfo(fiscal_year=self.target_year)
        # ลอง sheet ปี target ก่อน
        first_sheets = []
        for v in self._sheet_map.values():
            first_sheets.extend(v['sheets'])
        # ถ้าไม่มี ใช้ sheet แรก
        if not first_sheets:
            first_sheets = self.sheets[:3]

        for sname in first_sheets[:3]:
            df = self._read(sname)
            for _, row in df.iterrows():
                for j, val in enumerate(row):
                    if not isinstance(val, str):
                        continue
                    # ชื่อบริษัท
                    if any(kw in val for kw in self.COMPANY_KW) and not info.name:
                        nxt = [v for v in list(row.iloc[j+1:j+5])
                               if isinstance(v, str) and len(v) > 3
                               and not any(kw in v for kw in self.COMPANY_KW)]
                        if nxt:
                            info.name = nxt[0].strip()
                        elif any(biz in val for biz in ['หจก', 'บจก', 'บมจ', 'ห้างหุ้นส่วน', 'บริษัท']):
                            # ชื่อบริษัทอยู่ในเซลล์เดียวกัน
                            info.name = val.strip()
                    # เลขผู้เสียภาษี
                    if any(kw in val for kw in self.TAXID_KW) and not info.tax_id:
                        nxt = [v for v in list(row.iloc[j+1:j+5])
                               if str(v) not in ['nan', 'None']
                               and re.search(r'\d{10,13}', str(v))]
                        if nxt:
                            info.tax_id = re.sub(r'[^0-9]', '', str(nxt[0]))
            if info.name:
                break

        return info

    def _safe_float(self, val) -> float:
        """แปลงค่าเป็น float ปลอดภัย"""
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            return float(val) if not np.isnan(float(val)) else 0.0
        if isinstance(val, str):
            cleaned = re.sub(r'[,\s]', '', val)
            try:
                return float(cleaned)
            except:
                return 0.0
        return 0.0

    def _find_in_row(self, row, col_start: int, count: int = 5) -> list:
        """หาค่าตัวเลขในแถวถัดจาก col_start"""
        result = []
        for v in list(row.iloc[col_start:col_start+count]):
            if str(v) not in ['nan', 'None', ''] and isinstance(v, (int, float)):
                result.append(float(v))
        return result

    def _extract_from_sheet(self, df: pd.DataFrame) -> dict:
        """
        สกัดข้อมูลจาก DataFrame ของ 1 sheet
        พยายาม detect column layout อัตโนมัติ
        """
        data = {
            'sale_vat': 0., 'sale_exempt': 0., 'sale_zero': 0.,
            'purchase': 0., 'purchase_exempt': 0.,
            'vat_sale': 0., 'vat_buy': 0., 'vat_net': 0.,
            'gross_profit_rate': 0., 'confidence': 1.0
        }

        result_direct = self._extract_direct_layout(df)
        if result_direct['sale_vat'] > 0 or result_direct['purchase'] > 0:
            return result_direct

        # === Strategy 0: Scan ทุก col หา keyword+value pattern ===
        # รองรับ layout ที่ keyword และ value อยู่คนละ col
        keyword_col, value_col = self._find_keyword_value_cols(df)
        if keyword_col is not None:
            data.update(self._extract_by_keyword_col(df, keyword_col, value_col))
            if data['sale_vat'] > 0 or data['vat_net'] != 0:
                return data

        # === Strategy 1: หา row รวม แล้วอ่าน columns ===
        col_map = self._detect_column_layout(df)
        if col_map:
            data.update(self._read_by_col_map(df, col_map))
            return data

        # === Strategy 2: หา keyword rows ===
        for _, row in df.iterrows():
            row_str = ' '.join([str(v) for v in row if str(v) not in ['nan','None','']])

            # รวมรายได้ / ยอดขาย
            if any(kw in row_str for kw in self.TOTAL_SALE_KW):
                nums = [self._safe_float(v) for v in row
                        if isinstance(v,(int,float)) and self._safe_float(v) > 100]
                if len(nums) >= 2:
                    data['sale_vat']  = nums[0]
                    data['vat_sale']  = nums[1]
                elif len(nums) == 1:
                    data['sale_vat'] = nums[0]

            # ยอดขายยกเว้น
            if any(kw in row_str for kw in self.EXEMPT_KW):
                nums = [self._safe_float(v) for v in row
                        if isinstance(v,(int,float)) and self._safe_float(v) > 0]
                if nums:
                    data['sale_exempt'] = nums[0]

            # รวมค่าใช้จ่าย / ยอดซื้อ
            if any(kw in row_str for kw in self.TOTAL_BUY_KW):
                nums = [self._safe_float(v) for v in row
                        if isinstance(v,(int,float)) and self._safe_float(v) > 100]
                if len(nums) >= 2:
                    data['purchase'] = nums[0]
                    data['vat_buy']  = nums[1]
                elif len(nums) == 1:
                    data['purchase'] = nums[0]

            # ภาษีที่ต้องชำระ
            if any(kw in row_str for kw in self.VAT_PAY_KW):
                # หาเลขถัดจาก keyword
                for j, v in enumerate(row):
                    if isinstance(v, str) and any(kw in v for kw in self.VAT_PAY_KW):
                        nxt = self._find_in_row(row, j+1, 6)
                        if nxt:
                            data['vat_net'] = nxt[0]
                        break

            # อัตรากำไร
            if any(kw in row_str for kw in self.PROFIT_RATE_KW):
                for j, v in enumerate(row):
                    if isinstance(v, str) and any(kw in v for kw in self.PROFIT_RATE_KW):
                        nxt = self._find_in_row(row, j+1, 5)
                        if nxt:
                            data['gross_profit_rate'] = nxt[0]
                        break

        # === Strategy 3: ถ้ายังไม่ได้ vat_net คำนวณจาก vat_sale - vat_buy ===
        if data['vat_net'] == 0 and (data['vat_sale'] > 0 or data['vat_buy'] > 0):
            data['vat_net'] = data['vat_sale'] - data['vat_buy']
            data['confidence'] = 0.8  # ลด confidence เพราะ estimate

        # === ตรวจสอบ sanity ===
        data = self._sanity_check(data)
        return data

    def _extract_direct_layout(self, df):
        data = {'sale_vat':0.,'sale_exempt':0.,'purchase':0.,
                'vat_sale':0.,'vat_buy':0.,'vat_net':0.,
                'gross_profit_rate':0.,'confidence':1.0}
        try:
            v = df.iloc[2, 5]
            if isinstance(v, str) and len(v) > 3:
                self._last_company_name = v.strip()
        except: pass
        sale_found = purch_found = False
        for _, row in df.iterrows():
            if len(row) < 10: continue
            c1  = row.iloc[1]  if len(row)>1  else None
            v7  = row.iloc[7]  if len(row)>7  else None
            v8  = row.iloc[8]  if len(row)>8  else None
            v9  = row.iloc[9]  if len(row)>9  else None
            v10 = row.iloc[10] if len(row)>10 else None
            if (not sale_found and str(c1) in ['nan','None',''] and
                isinstance(v7,(int,float)) and v7 > 50000 and
                isinstance(v8,(int,float)) and isinstance(v9,(int,float)) and
                abs(v7-(v8+v9)) < 5):
                data['sale_vat'] = float(v9); data['vat_sale'] = float(v8)
                sale_found = True
            if (v10 == 402 and isinstance(v7,(int,float)) and
                (str(v8) in ['nan','None',''] or v8 == 0)):
                data['sale_exempt'] += float(v7)
            if not purch_found and len(row) > 32:
                c24 = row.iloc[24] if len(row)>24 else None
                v30 = row.iloc[30] if len(row)>30 else None
                v31 = row.iloc[31] if len(row)>31 else None
                v32 = row.iloc[32] if len(row)>32 else None
                if (str(c24) in ['nan','None',''] and
                    isinstance(v30,(int,float)) and v30 > 10000 and
                    isinstance(v31,(int,float)) and isinstance(v32,(int,float)) and
                    abs(v30-(v31+v32)) < 5):
                    data['purchase'] = float(v32); data['vat_buy'] = float(v31)
                    purch_found = True
            for j, v in enumerate(row):
                if isinstance(v, str) and 'รวมชำระ ภ.พ.30' in v:
                    nxt=[x for x in list(row.iloc[j+1:j+5])
                         if str(x) not in ['nan','None'] and isinstance(x,(int,float))]
                    if nxt: data['vat_net'] = float(nxt[0])
        return data

    def _find_keyword_value_cols(self, df: pd.DataFrame):
        """
        ค้นหา column ที่เป็น keyword และ column ที่เป็นค่าตัวเลขคู่กัน
        รองรับ layout ที่ keyword อยู่ col ไหนก็ได้
        """
        ALL_KW = (self.TOTAL_SALE_KW + self.TOTAL_BUY_KW +
                  self.VAT_PAY_KW + self.EXEMPT_KW)
        found = {}
        for _, row in df.iterrows():
            for j, v in enumerate(row):
                if isinstance(v, str) and any(kw in v for kw in ALL_KW):
                    found[j] = found.get(j, 0) + 1
        if not found:
            return None, None
        kw_col = max(found, key=found.get)
        num_count = {}
        for _, row in df.iterrows():
            for j in range(kw_col + 1, min(kw_col + 6, len(row))):
                v = row.iloc[j] if j < len(row) else None
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    if abs(float(v)) > 100:
                        num_count[j] = num_count.get(j, 0) + 1
        if not num_count:
            return kw_col, kw_col + 1
        val_col = max(num_count, key=num_count.get)
        return kw_col, val_col

    def _extract_by_keyword_col(self, df: pd.DataFrame,
                                kw_col: int, val_col: int) -> dict:
        """สกัดข้อมูลโดยใช้ keyword column และ value column ที่ detect ได้"""
        data = {'sale_vat': 0., 'sale_exempt': 0., 'purchase': 0.,
                'vat_sale': 0., 'vat_buy': 0., 'vat_net': 0.,
                'gross_profit_rate': 0.}
        for _, row in df.iterrows():
            if kw_col >= len(row):
                continue
            kw = str(row.iloc[kw_col]).strip()
            val = 0.
            for vc in range(val_col, min(val_col + 4, len(row))):
                v = row.iloc[vc]
                if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(float(v)) > 0:
                    val = float(v); break

            if any(k in kw for k in self.TOTAL_SALE_KW):
                data['sale_vat'] = val
                for vc2 in range(val_col + 1, min(val_col + 5, len(row))):
                    v2 = row.iloc[vc2]
                    if isinstance(v2, (int, float)) and abs(float(v2)) > 0:
                        data['vat_sale'] = float(v2); break
            elif any(k in kw for k in self.EXEMPT_KW):
                data['sale_exempt'] = val
            elif any(k in kw for k in self.TOTAL_BUY_KW):
                data['purchase'] = val
                for vc2 in range(val_col + 1, min(val_col + 5, len(row))):
                    v2 = row.iloc[vc2]
                    if isinstance(v2, (int, float)) and abs(float(v2)) > 0:
                        data['vat_buy'] = float(v2); break
            elif any(k in kw for k in self.VAT_PAY_KW):
                if data['vat_net'] == 0:
                    data['vat_net'] = val
            elif any(k in kw for k in self.PROFIT_RATE_KW):
                data['gross_profit_rate'] = val
        return data

    def _detect_column_layout(self, df: pd.DataFrame) -> dict:
        """
        Detect layout ของ columns จาก header rows
        คืน dict: {'sale_col': N, 'vat_sale_col': M, ...}
        """
        col_map = {}
        # หา header row (มักอยู่ใน row 1-10)
        for i in range(min(15, len(df))):
            row = df.iloc[i]
            row_vals = [str(v).strip() for v in row]
            row_str = ' '.join(row_vals)

            # detect header ฝั่งขาย
            if 'มูลค่า' in row_str and ('ขาย' in row_str or 'รายได้' in row_str):
                for j, v in enumerate(row_vals):
                    if 'มูลค่า' in v and ('ขาย' in v or 'รายได้' in v):
                        col_map['sale_col'] = j
                    if 'ภาษีมูลค่าเพิ่ม' in v or 'ภาษีขาย' in v:
                        col_map['vat_sale_col'] = j
                    if 'มูลค่า' in v and ('ซื้อ' in v or 'ต้นทุน' in v):
                        col_map['purchase_col'] = j
                    if 'ภาษีซื้อ' in v:
                        col_map['vat_buy_col'] = j

            # ถ้าเจอ 'รวม' ใน col แรกๆ น่าจะเป็น total row
            if row_vals[0] == 'รวม' and len(col_map) > 0:
                col_map['total_row'] = i

        return col_map if len(col_map) >= 2 else {}

    def _read_by_col_map(self, df: pd.DataFrame, col_map: dict) -> dict:
        """อ่านตัวเลขตาม column map ที่ detect ได้"""
        data = {}
        if 'total_row' not in col_map:
            return data
        row = df.iloc[col_map['total_row']]
        if 'sale_col' in col_map:
            data['sale_vat'] = self._safe_float(row.iloc[col_map['sale_col']])
        if 'vat_sale_col' in col_map:
            data['vat_sale'] = self._safe_float(row.iloc[col_map['vat_sale_col']])
        if 'purchase_col' in col_map:
            data['purchase'] = self._safe_float(row.iloc[col_map['purchase_col']])
        if 'vat_buy_col' in col_map:
            data['vat_buy'] = self._safe_float(row.iloc[col_map['vat_buy_col']])
        return data

    def _sanity_check(self, data: dict) -> dict:
        """ตรวจสอบความสมเหตุสมผลของตัวเลข"""
        # ภาษีขายต้องประมาณ 7% ของยอดขาย
        if data['sale_vat'] > 0 and data['vat_sale'] > 0:
            ratio = data['vat_sale'] / data['sale_vat']
            if not (0.05 <= ratio <= 0.09):  # ยอมรับ 5-9%
                data['confidence'] = min(data['confidence'], 0.7)
                log.warning(f"อัตราภาษีขาย {ratio:.2%} ผิดปกติ (ควร ~7%)")

        # ภาษีซื้อต้องประมาณ 7% ของยอดซื้อ
        if data['purchase'] > 0 and data['vat_buy'] > 0:
            ratio = data['vat_buy'] / data['purchase']
            if not (0.05 <= ratio <= 0.09):
                data['confidence'] = min(data['confidence'], 0.7)
                log.warning(f"อัตราภาษีซื้อ {ratio:.2%} ผิดปกติ (ควร ~7%)")

        return data

    def extract_month(self, month_num: int) -> MonthlyVAT:
        """สกัดข้อมูลของ 1 เดือน"""
        mv = MonthlyVAT(
            month_num=month_num,
            month_th=self.MONTHS_TH[month_num-1],
            month_full=self.MONTHS_FULL[month_num-1]
        )

        if month_num not in self._sheet_map:
            mv.notes = "ไม่พบ sheet สำหรับเดือนนี้"
            mv.confidence = 0.0
            return mv

        info = self._sheet_map[month_num]
        mv.sheet_type = info['type']
        mv.source_sheet = ', '.join(info['sheets'])

        # รวมข้อมูลจากทุก sheet ของเดือนนี้
        combined = {
            'sale_vat': 0., 'sale_exempt': 0., 'sale_zero': 0.,
            'purchase': 0., 'purchase_exempt': 0.,
            'vat_sale': 0., 'vat_buy': 0., 'vat_net': 0.,
            'gross_profit_rate': 0., 'confidence': 1.0
        }

        for sname in info['sheets']:
            df = self._read(sname)
            extracted = self._extract_from_sheet(df)

            if info['type'] == 'summary':
                # sheet สรุปมีข้อมูลครบ ใช้เลยไม่ต้อง +
                combined = extracted
                break
            else:
                # pair หรือ single: บวกรวม
                for key in ['sale_vat','sale_exempt','purchase','vat_sale','vat_buy']:
                    combined[key] += extracted.get(key, 0.)
                combined['vat_net'] += extracted.get('vat_net', 0.)
                combined['confidence'] = min(combined['confidence'],
                                             extracted.get('confidence', 1.0))
                if extracted.get('gross_profit_rate', 0.) > 0:
                    combined['gross_profit_rate'] = extracted['gross_profit_rate']

        # map ลงใน MonthlyVAT
        mv.sale_vat        = combined['sale_vat']
        mv.sale_exempt     = combined['sale_exempt']
        mv.sale_zero       = combined.get('sale_zero', 0.)
        mv.total_sale      = combined['sale_vat'] + combined['sale_exempt'] + combined.get('sale_zero', 0.)
        mv.purchase        = combined['purchase']
        mv.total_purchase  = combined['purchase'] + combined.get('purchase_exempt', 0.)
        mv.vat_sale        = combined['vat_sale']
        mv.vat_buy         = combined['vat_buy']
        mv.vat_net         = combined['vat_net']
        mv.gross_profit    = mv.total_sale - mv.purchase
        mv.gross_profit_rate = combined['gross_profit_rate']
        mv.confidence      = combined['confidence']

        return mv

    def extract_all(self) -> list:
        """สกัดข้อมูลทุกเดือน พร้อมคำนวณ vat_cumul"""
        monthly = []
        cumul = 0.
        for m in range(1, 13):
            mv = self.extract_month(m)
            cumul += mv.vat_net
            mv.vat_cumul = cumul
            monthly.append(mv)
            log.info(f"  {mv.month_th}: ขาย={mv.sale_vat:,.0f} "
                     f"ซื้อ={mv.purchase:,.0f} "
                     f"VAT={mv.vat_net:,.2f} "
                     f"[conf={mv.confidence:.0%}]")
        return monthly


# ══════════════════════════════════════════════════════════════
# 2. AccountClassifier — คัดแยกและจัดหมวดบัญชี
# ══════════════════════════════════════════════════════════════

class AccountClassifier:
    """
    คัดแยกและจัดหมวดบัญชีจากข้อมูล ภพ.30
    แยกประเภทบัญชีอัตโนมัติตาม Thai accounting standards
    """

    # หมวดบัญชีตามมาตรฐานการบัญชีไทย
    ACCOUNT_TYPES = {
        '1': 'สินทรัพย์',
        '11': 'สินทรัพย์หมุนเวียน',
        '12': 'สินทรัพย์ไม่หมุนเวียน',
        '2': 'หนี้สิน',
        '21': 'หนี้สินหมุนเวียน',
        '22': 'หนี้สินไม่หมุนเวียน',
        '3': 'ส่วนของเจ้าของ',
        '4': 'รายได้',
        '41': 'รายได้จากการขาย',
        '42': 'รายได้อื่น',
        '5': 'ค่าใช้จ่าย',
        '51': 'ต้นทุนขาย',
        '52': 'ค่าใช้จ่ายในการขาย',
        '53': 'ค่าใช้จ่ายในการบริหาร',
        '54': 'ต้นทุนทางการเงิน',
    }

    # ประเภทธุรกิจ → การตีความยอดขาย
    BUSINESS_TYPES = {
        'ค้าปลีก':     {'vat_rate': 0.07, 'margin_range': (0.10, 0.40)},
        'ค้าส่ง':      {'vat_rate': 0.07, 'margin_range': (0.05, 0.20)},
        'บริการ':      {'vat_rate': 0.07, 'margin_range': (0.30, 0.80)},
        'ก่อสร้าง':    {'vat_rate': 0.07, 'margin_range': (0.10, 0.30)},
        'อสังหาริมทรัพย์': {'vat_rate': 0.07, 'margin_range': (0.15, 0.50)},
        'รถยนต์':      {'vat_rate': 0.07, 'margin_range': (0.05, 0.20)},
        'อาหาร':       {'vat_rate': 0.07, 'margin_range': (0.20, 0.60)},
        'ทั่วไป':      {'vat_rate': 0.07, 'margin_range': (0.08, 0.50)},
    }

    def __init__(self, business_type: str = 'ทั่วไป'):
        self.business_type = business_type
        self.biz_config = self.BUSINESS_TYPES.get(business_type,
                          self.BUSINESS_TYPES['ทั่วไป'])

    def detect_business_type(self, company_info: CompanyInfo,
                              monthly: list) -> str:
        """ตรวจสอบประเภทธุรกิจจากชื่อบริษัทและรูปแบบรายได้"""
        name_lower = company_info.name.lower()
        biz_type = 'ทั่วไป'

        keywords_map = {
            'รถ': 'รถยนต์', 'ยานยนต์': 'รถยนต์', 'auto': 'รถยนต์',
            'ก่อสร้าง': 'ก่อสร้าง', 'construction': 'ก่อสร้าง',
            'อาหาร': 'อาหาร', 'food': 'อาหาร', 'ร้านอาหาร': 'อาหาร',
            'อสังหา': 'อสังหาริมทรัพย์', 'property': 'อสังหาริมทรัพย์',
            'บริการ': 'บริการ', 'service': 'บริการ',
            'ค้าส่ง': 'ค้าส่ง', 'wholesale': 'ค้าส่ง',
            'ค้าปลีก': 'ค้าปลีก', 'retail': 'ค้าปลีก',
        }
        for kw, btype in keywords_map.items():
            if kw in name_lower:
                biz_type = btype
                break

        # ตรวจสอบจาก gross margin ถ้ามีข้อมูล
        valid = [m for m in monthly if m.total_sale > 0 and m.gross_profit > 0]
        if valid:
            avg_margin = sum(m.gross_profit/m.total_sale for m in valid) / len(valid)
            # ปรับ business type ตาม margin
            if avg_margin > 0.5:
                biz_type = 'บริการ'
            elif avg_margin < 0.08:
                biz_type = 'ค้าส่ง'

        self.business_type = biz_type
        self.biz_config = self.BUSINESS_TYPES.get(biz_type,
                          self.BUSINESS_TYPES['ทั่วไป'])
        log.info(f"ประเภทธุรกิจที่ตรวจพบ: {biz_type}")
        return biz_type

    def classify_vat_items(self, monthly: list) -> dict:
        """
        แยกประเภทรายการจาก ภพ.30:
        - รายได้ที่มี VAT vs ยกเว้น VAT
        - ซื้อในประเทศ vs นำเข้า
        - รายการปกติ vs ผิดปกติ
        """
        total_sale_vat    = sum(m.sale_vat     for m in monthly)
        total_sale_exempt = sum(m.sale_exempt   for m in monthly)
        total_sale        = sum(m.total_sale    for m in monthly)
        total_purchase    = sum(m.purchase      for m in monthly)
        total_vat_sale    = sum(m.vat_sale      for m in monthly)
        total_vat_buy     = sum(m.vat_buy       for m in monthly)
        total_vat_net     = sum(m.vat_net       for m in monthly)

        # คำนวณอัตราภาษีจริง
        effective_vat_rate_sale = (total_vat_sale / total_sale_vat
                                    if total_sale_vat > 0 else 0)
        effective_vat_rate_buy  = (total_vat_buy / total_purchase
                                    if total_purchase > 0 else 0)

        # gross margin
        avg_margin = ((total_sale - total_purchase) / total_sale
                      if total_sale > 0 else 0)

        return {
            'revenue_breakdown': {
                'vat_sales':    total_sale_vat,
                'exempt_sales': total_sale_exempt,
                'total_sales':  total_sale,
                'pct_exempt':   total_sale_exempt/total_sale if total_sale > 0 else 0,
            },
            'purchase_breakdown': {
                'total_purchase':     total_purchase,
                'est_vat_purchase':   total_vat_buy * 107/7,  # ประมาณการ
            },
            'tax_analysis': {
                'total_vat_sale':         total_vat_sale,
                'total_vat_buy':          total_vat_buy,
                'total_vat_net':          total_vat_net,
                'effective_vat_rate_sale': effective_vat_rate_sale,
                'effective_vat_rate_buy':  effective_vat_rate_buy,
                'expected_vat_rate':       0.07,
                'vat_rate_ok':            (0.05 <= effective_vat_rate_sale <= 0.09
                                           if total_sale_vat > 0 else True),
            },
            'profitability': {
                'gross_profit':  total_sale - total_purchase,
                'gross_margin':  avg_margin,
                'margin_range':  self.biz_config['margin_range'],
                'margin_ok':     (self.biz_config['margin_range'][0]
                                  <= avg_margin
                                  <= self.biz_config['margin_range'][1]),
            },
        }

    def detect_anomalies(self, monthly: list) -> list:
        """ตรวจหาความผิดปกติในข้อมูลรายเดือน"""
        anomalies = []
        sales = [m.total_sale for m in monthly if m.total_sale > 0]
        if not sales:
            return anomalies

        mean_sale = np.mean(sales)
        std_sale  = np.std(sales)

        for m in monthly:
            if m.total_sale == 0:
                continue

            # 1. ยอดขายสูง/ต่ำผิดปกติ (> 2.5 SD)
            if std_sale > 0 and abs(m.total_sale - mean_sale) > 2.5 * std_sale:
                direction = "สูง" if m.total_sale > mean_sale else "ต่ำ"
                anomalies.append({
                    'month': m.month_th,
                    'type': 'ยอดขายผิดปกติ',
                    'detail': f"ยอดขาย {m.total_sale:,.0f} {direction}กว่าค่าเฉลี่ย"
                              f" {mean_sale:,.0f} มากผิดปกติ",
                    'severity': 'high'
                })

            # 2. อัตราภาษีผิดปกติ
            if m.sale_vat > 0 and m.vat_sale > 0:
                rate = m.vat_sale / m.sale_vat
                if not (0.04 <= rate <= 0.10):
                    anomalies.append({
                        'month': m.month_th,
                        'type': 'อัตราภาษีผิดปกติ',
                        'detail': f"อัตราภาษีขาย {rate:.2%} ไม่ใกล้เคียง 7%",
                        'severity': 'medium'
                    })

            # 3. confidence ต่ำ
            if m.confidence < 0.7:
                anomalies.append({
                    'month': m.month_th,
                    'type': 'ข้อมูลไม่ชัดเจน',
                    'detail': f"ความมั่นใจในข้อมูลต่ำ ({m.confidence:.0%})"
                              " — ควรตรวจสอบ sheet ต้นทาง",
                    'severity': 'high'
                })

            # 4. ยอดซื้อมากกว่ายอดขายมาก (ไม่สมเหตุสมผล)
            if (m.purchase > 0 and m.total_sale > 0
                    and m.purchase > m.total_sale * 1.5):
                anomalies.append({
                    'month': m.month_th,
                    'type': 'ยอดซื้อสูงกว่ายอดขายมาก',
                    'detail': f"ซื้อ {m.purchase:,.0f} > ขาย {m.total_sale:,.0f}",
                    'severity': 'medium'
                })

        return anomalies


# ══════════════════════════════════════════════════════════════
# 3. FinancialCalc — คำนวณงบการเงิน
# ══════════════════════════════════════════════════════════════

class FinancialCalc:
    """คำนวณงบการเงินจากข้อมูล ภพ.30"""

    # อัตราภาษีเงินได้นิติบุคคล (ปี 2568)
    CIT_RATES = [
        (300_000,   0.00),   # 0-300K: 0%
        (3_000_000, 0.15),   # 300K-3M: 15%
        (float('inf'), 0.20) # >3M: 20%
    ]

    def __init__(self,
                 retained_begin: float = 0.,
                 admin_expense: float = 0.,
                 other_income: float = 0.,
                 other_expense: float = 0.,
                 inventory_begin: float = 0.,
                 inventory_end: float = 0.):
        self.retained_begin  = retained_begin
        self.admin_expense   = admin_expense
        self.other_income    = other_income
        self.other_expense   = other_expense
        self.inventory_begin = inventory_begin
        self.inventory_end   = inventory_end

    def calc_cit(self, taxable_profit: float) -> float:
        """คำนวณภาษีเงินได้นิติบุคคลตามอัตราก้าวหน้า"""
        if taxable_profit <= 0:
            return 0.
        tax = 0.
        prev_limit = 0.
        for limit, rate in self.CIT_RATES:
            bracket = min(taxable_profit, limit) - prev_limit
            if bracket <= 0:
                break
            tax += bracket * rate
            prev_limit = limit
        return tax

    def calculate(self, monthly: list) -> dict:
        """คำนวณงบการเงินทั้งปี"""
        valid = [m for m in monthly if m.total_sale > 0]

        total_sale      = sum(m.total_sale   for m in valid)
        total_purchase  = sum(m.purchase     for m in valid)
        total_vat_sale  = sum(m.vat_sale     for m in valid)
        total_vat_buy   = sum(m.vat_buy      for m in valid)
        total_vat_net   = sum(m.vat_net      for m in valid)

        # คำนวณ COGS (ใช้ข้อมูลสินค้าคงเหลือถ้ามี)
        cogs = (self.inventory_begin + total_purchase - self.inventory_end
                if self.inventory_end > 0
                else total_purchase)

        # PL
        gross_profit    = total_sale + self.other_income - cogs
        ebit            = gross_profit - self.admin_expense - self.other_expense
        taxable_profit  = max(ebit, 0.)
        cit             = self.calc_cit(taxable_profit)
        net_profit      = ebit - cit
        retained_end    = self.retained_begin + net_profit

        return {
            'pl': {
                'revenue':         total_sale,
                'other_income':    self.other_income,
                'total_revenue':   total_sale + self.other_income,
                'cogs':            cogs,
                'gross_profit':    total_sale - cogs,
                'gross_margin':    (total_sale-cogs)/total_sale if total_sale else 0,
                'admin_expense':   self.admin_expense,
                'other_expense':   self.other_expense,
                'total_expense':   cogs + self.admin_expense + self.other_expense,
                'ebit':            ebit,
                'finance_cost':    0.,
                'ebt':             ebit,
                'cit':             -cit,
                'net_profit':      net_profit,
                'retained_begin':  self.retained_begin,
                'retained_end':    retained_end,
            },
            'vat': {
                'total_sale_vat':  sum(m.sale_vat    for m in valid),
                'total_exempt':    sum(m.sale_exempt  for m in valid),
                'total_sale':      total_sale,
                'total_purchase':  total_purchase,
                'total_vat_sale':  total_vat_sale,
                'total_vat_buy':   total_vat_buy,
                'total_vat_net':   total_vat_net,
            },
            'tax_payable': {
                'vat':      total_vat_net,
                'cit':      -cit,    # ลบ = ชำระเกิน/ขอคืน
                'taxable_profit': taxable_profit,
            },
        }


# ══════════════════════════════════════════════════════════════
# 4. GeminiAnalyzer — วิเคราะห์ด้วย Gemini
# ══════════════════════════════════════════════════════════════

class GeminiAnalyzer:
    """ส่งข้อมูลให้ Gemini วิเคราะห์และตรวจสอบก่อนปิดงบ"""

    SYSTEM_PROMPT = """
คุณคือนักบัญชีผู้เชี่ยวชาญด้านภาษีและการปิดงบการเงินสำหรับธุรกิจ SME ไทย
มีความเชี่ยวชาญ:
- ภาษีมูลค่าเพิ่ม (VAT) และ ภ.พ.30
- ภาษีเงินได้นิติบุคคล (CIT) และ ภงด.50/51
- มาตรฐานการบัญชีไทย (NPAEs)
- การตรวจสอบความสมเหตุสมผลของงบการเงิน

วิเคราะห์ข้อมูลด้วยความรอบคอบ ระบุจุดเสี่ยง และให้คำแนะนำที่ปฏิบัติได้จริง
ตอบเป็นภาษาไทย กระชับ และตรงประเด็น
"""

    def __init__(self, api_key: str, model_name: str = 'gemini-2.5-flash',
                 max_retries: int = 3):
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(model_name)
        self.max_retries = max_retries

    def _call_api(self, prompt: str) -> str:
        """เรียก Gemini API พร้อม retry"""
        for attempt in range(self.max_retries):
            try:
                resp = self.model.generate_content(
                    f"{self.SYSTEM_PROMPT}\n\n{prompt}"
                )
                return resp.text
            except Exception as e:
                if '429' in str(e) and attempt < self.max_retries - 1:
                    wait = 15 * (attempt + 1)
                    log.warning(f"Rate limit — รอ {wait} วิ...")
                    time.sleep(wait)
                else:
                    log.error(f"Gemini Error: {e}")
                    return f"ERROR: {e}"
        return "ERROR: เรียก API ไม่สำเร็จ"

    def analyze_vat(self, company: CompanyInfo, monthly: list,
                    classification: dict, anomalies: list,
                    calc_result: dict) -> tuple:
        """วิเคราะห์ข้อมูล ภพ.30 และงบการเงิน"""

        # สร้างสรุปรายเดือน
        month_lines = []
        for m in monthly:
            if m.total_sale > 0:
                line = (f"- {m.month_full}: ขาย {m.sale_vat:,.0f}"
                        f" ยกเว้น {m.sale_exempt:,.0f}"
                        f" | ซื้อ {m.purchase:,.0f}"
                        f" | ภาษีขาย {m.vat_sale:,.0f}"
                        f" ภาษีซื้อ {m.vat_buy:,.0f}"
                        f" | ชำระ {m.vat_net:,.2f}"
                        f" | margin {m.gross_profit_rate:.1%}"
                        f" [conf={m.confidence:.0%}]")
                month_lines.append(line)

        vat   = classification['tax_analysis']
        profit = classification['profitability']
        pl     = calc_result['pl']
        anomaly_text = "\n".join([
            f"  ⚠ [{a['severity'].upper()}] {a['month']}: {a['type']} — {a['detail']}"
            for a in anomalies
        ]) or "  (ไม่พบความผิดปกติ)"

        prompt = f"""
วิเคราะห์ข้อมูลงบการเงิน ภพ.30 ต่อไปนี้:

┌─ ข้อมูลบริษัท ─────────────────────────────┐
│ ชื่อ: {company.name}
│ เลขผู้เสียภาษี: {company.tax_id}
│ ปีภาษี: 25{company.fiscal_year}
└────────────────────────────────────────────┘

┌─ ข้อมูล ภพ.30 รายเดือน ───────────────────┐
{chr(10).join(month_lines)}
└────────────────────────────────────────────┘

┌─ สรุปรวมทั้งปี ────────────────────────────┐
│ รายได้รวม:        {pl['total_revenue']:>15,.2f} บาท
│ ต้นทุนขาย:        {pl['cogs']:>15,.2f} บาท
│ กำไรขั้นต้น:      {pl['gross_profit']:>15,.2f} บาท ({profit['gross_margin']:.1%})
│ ค่าใช้จ่ายบริหาร: {pl['admin_expense']:>15,.2f} บาท
│ EBIT:             {pl['ebit']:>15,.2f} บาท
│ ภาษีเงินได้:      {pl['cit']:>15,.2f} บาท
│ กำไรสุทธิ:        {pl['net_profit']:>15,.2f} บาท
├────────────────────────────────────────────┤
│ ภาษีขายรวม:      {vat['total_vat_sale']:>15,.2f} บาท (rate={vat['effective_vat_rate_sale']:.2%})
│ ภาษีซื้อรวม:     {vat['total_vat_buy']:>15,.2f} บาท (rate={vat['effective_vat_rate_buy']:.2%})
│ ภาษี VAT สุทธิ:  {vat['total_vat_net']:>15,.2f} บาท
└────────────────────────────────────────────┘

┌─ ความผิดปกติที่ตรวจพบ ─────────────────────┐
{anomaly_text}
└────────────────────────────────────────────┘

กรุณาวิเคราะห์ 5 ด้านต่อไปนี้:

1. **ความสมเหตุสมผลรายเดือน**
   - เดือนไหนผิดปกติชัดเจน? เพราะอะไร?
   - แนวโน้มรายได้และต้นทุนสอดคล้องกันไหม?

2. **ความถูกต้องของภาษี VAT**
   - อัตราภาษีขาย/ซื้อสมเหตุสมผลไหม (ควร ~7%)?
   - มีเดือนที่ภาษีดูผิดปกติไหม?

3. **อัตรากำไรและต้นทุน**
   - Gross margin {profit['gross_margin']:.1%} เหมาะสมกับประเภทธุรกิจไหม?
   - ต้นทุนขายสัมพันธ์กับยอดขายอย่างสมเหตุสมผลไหม?

4. **จุดเสี่ยงสำคัญ**
   - มีประเด็นที่อาจมีปัญหากับกรมสรรพากรไหม?
   - มีรายการที่ควรขอเอกสารเพิ่มเติมไหม?

5. **สรุปและคำแนะนำ**
   - ข้อมูลพร้อมปิดงบหรือยัง?
   - ถ้ามีปัญหา ควรแก้ไขอะไรก่อน?

ท้ายสุดระบุ:
VERDICT: พร้อมปิดงบ
หรือ
VERDICT: ต้องตรวจสอบเพิ่มเติม — [ระบุสิ่งที่ต้องแก้]
"""
        result = self._call_api(prompt)
        verdict = ('ready' if 'VERDICT: พร้อมปิดงบ' in result
                   else 'review' if 'VERDICT:' in result
                   else 'error')
        return result, verdict

    def suggest_adjustments(self, company: CompanyInfo,
                             calc_result: dict,
                             anomalies: list) -> str:
        """แนะนำรายการปรับปรุงบัญชี (Adjusting Entries)"""
        prompt = f"""
จากข้อมูลงบการเงินของ {company.name} ปี 25{company.fiscal_year}:

กำไรก่อนภาษี: {calc_result['pl']['ebt']:,.2f} บาท
ภาษีเงินได้: {calc_result['pl']['cit']:,.2f} บาท

ความผิดปกติ:
{json.dumps([a['detail'] for a in anomalies], ensure_ascii=False, indent=2)}

กรุณาแนะนำ:
1. รายการปรับปรุงบัญชีที่ควรพิจารณา (ถ้ามี)
2. รายการค้างรับ/ค้างจ่ายที่ควรบันทึก
3. การตรวจสอบสต็อกสินค้า (ถ้าเกี่ยวข้อง)
4. ประเด็นภาษีที่ควรระวัง

ตอบสั้นๆ เป็นข้อๆ
"""
        return self._call_api(prompt)


# ══════════════════════════════════════════════════════════════
# 5. FinancialEngine — Orchestrator หลัก
# ══════════════════════════════════════════════════════════════

class FinancialEngine:
    """
    Orchestrator รวม VAT30Reader + AccountClassifier
    + FinancialCalc + GeminiAnalyzer

    วิธีใช้:
        engine = FinancialEngine(
            gemini_api_key="...",
            target_year="68",
            retained_begin=2278127.06,
            admin_expense=5000,
        )
        result = engine.process("ภพ.30_บริษัทXYZ.xlsx")
        result.export("output.xlsx")
    """

    def __init__(self,
                 gemini_api_key: str,
                 target_year: str = "68",
                 gemini_model: str = "gemini-2.5-flash",
                 # FinancialCalc params (ปรับตามลูกค้า)
                 retained_begin: float = 0.,
                 admin_expense: float = 0.,
                 other_income: float = 0.,
                 inventory_begin: float = 0.,
                 inventory_end: float = 0.,
                 business_type: str = 'ทั่วไป'):

        self.target_year    = target_year
        self.reader         = VAT30Reader(target_year)
        self.classifier     = AccountClassifier(business_type)
        self.calc           = FinancialCalc(
            retained_begin=retained_begin,
            admin_expense=admin_expense,
            other_income=other_income,
            inventory_begin=inventory_begin,
            inventory_end=inventory_end,
        )
        self.analyzer       = GeminiAnalyzer(gemini_api_key, gemini_model)
        self._summary       = None

    def process(self, filepath: str,
                skip_gemini: bool = False) -> FinancialSummary:
        """
        ประมวลผลไฟล์ ภพ.30 ครบทุกขั้นตอน
        Args:
            filepath: path ไปยังไฟล์ Excel
            skip_gemini: True = ข้ามการวิเคราะห์ด้วย AI (ทดสอบ)
        Returns:
            FinancialSummary ที่มีข้อมูลครบถ้วน
        """
        log.info(f"{'='*55}")
        log.info(f"เริ่มประมวลผล: {filepath}")
        log.info(f"{'='*55}")

        # Step 1: อ่านไฟล์
        log.info("Step 1: อ่านไฟล์ ภพ.30...")
        self.reader.load(filepath)
        company = self.reader.get_company_info()
        log.info(f"บริษัท: {company.name} | เลขผู้เสียภาษี: {company.tax_id}")

        # Step 2: สกัดข้อมูลรายเดือน
        log.info("\nStep 2: สกัดข้อมูลรายเดือน...")
        monthly = self.reader.extract_all()

        # Step 3: ตรวจสอบประเภทธุรกิจ
        log.info("\nStep 3: ตรวจสอบประเภทธุรกิจ...")
        biz_type = self.classifier.detect_business_type(company, monthly)

        # Step 4: จัดหมวดบัญชีและหาความผิดปกติ
        log.info("\nStep 4: วิเคราะห์บัญชีและหาความผิดปกติ...")
        classification = self.classifier.classify_vat_items(monthly)
        anomalies      = self.classifier.detect_anomalies(monthly)
        if anomalies:
            log.warning(f"พบความผิดปกติ {len(anomalies)} รายการ")
            for a in anomalies:
                log.warning(f"  [{a['severity']}] {a['month']}: {a['detail']}")

        # Step 5: คำนวณงบการเงิน
        log.info("\nStep 5: คำนวณงบการเงิน...")
        calc_result = self.calc.calculate(monthly)
        pl = calc_result['pl']
        log.info(f"  รายได้รวม:   {pl['total_revenue']:>15,.2f}")
        log.info(f"  กำไรขั้นต้น: {pl['gross_profit']:>15,.2f} ({pl['gross_margin']:.1%})")
        log.info(f"  กำไรสุทธิ:   {pl['net_profit']:>15,.2f}")

        # Step 6: Gemini วิเคราะห์
        analysis_text, verdict = "", "skip"
        if not skip_gemini:
            log.info("\nStep 6: Gemini วิเคราะห์...")
            analysis_text, verdict = self.analyzer.analyze_vat(
                company, monthly, classification, anomalies, calc_result
            )
            log.info(f"Gemini Verdict: {verdict}")
        else:
            log.info("\nStep 6: ข้าม Gemini (skip_gemini=True)")

        # รวมผลลัพธ์
        self._summary = FinancialSummary(
            company        = company,
            year           = self.target_year,
            monthly        = monthly,
            total_revenue  = pl['total_revenue'],
            total_cogs     = pl['cogs'],
            gross_profit   = pl['gross_profit'],
            admin_expense  = pl['admin_expense'],
            ebit           = pl['ebit'],
            tax_expense    = pl['cit'],
            net_profit     = pl['net_profit'],
            retained_begin = pl['retained_begin'],
            retained_end   = pl['retained_end'],
            total_vat_sale = calc_result['vat']['total_vat_sale'],
            total_vat_buy  = calc_result['vat']['total_vat_buy'],
            total_vat_net  = calc_result['vat']['total_vat_net'],
            gemini_analysis = analysis_text,
            gemini_verdict  = verdict,
            anomalies       = anomalies,
        )

        log.info(f"\n{'='*55}")
        log.info(f"ประมวลผลเสร็จสิ้น | Verdict: {verdict}")
        log.info(f"{'='*55}")
        return self._summary

    def export(self, summary: FinancialSummary,
               output_path: str = None) -> str:
        """Export ผลลัพธ์เป็น Excel (ใช้ Styles เดิม)"""
        from financial_exporter import FinancialExporter
        exporter = FinancialExporter()
        return exporter.export(summary, output_path)

    def print_summary(self, summary: FinancialSummary):
        """แสดงสรุปผลใน console"""
        print(f"\n{'═'*60}")
        print(f"  สรุปงบการเงิน: {summary.company.name}")
        print(f"  ปี พ.ศ. 25{summary.year}")
        print(f"{'═'*60}")
        print(f"\n  {'เดือน':^6} {'ยอดขาย':>14} {'ยอดซื้อ':>14} "
              f"{'VAT ชำระ':>10} {'Conf':>6}")
        print(f"  {'-'*54}")
        for m in summary.monthly:
            if m.total_sale > 0:
                conf_bar = '●' * int(m.confidence * 5)
                print(f"  {m.month_th:^6} {m.total_sale:>14,.0f} "
                      f"{m.purchase:>14,.0f} {m.vat_net:>10,.2f} "
                      f"{conf_bar:<5}")
        print(f"  {'-'*54}")
        valid = [m for m in summary.monthly if m.total_sale > 0]
        print(f"  {'รวม':^6} {sum(m.total_sale for m in valid):>14,.0f} "
              f"{sum(m.purchase for m in valid):>14,.0f} "
              f"{summary.total_vat_net:>10,.2f}")
        print(f"\n  รายได้รวม:    {summary.total_revenue:>14,.2f} บาท")
        print(f"  กำไรสุทธิ:    {summary.net_profit:>14,.2f} บาท")
        print(f"  ภาษี VAT สุทธิ:{summary.total_vat_net:>13,.2f} บาท")
        if summary.anomalies:
            print(f"\n  ⚠ ความผิดปกติ {len(summary.anomalies)} รายการ:")
            for a in summary.anomalies:
                print(f"    [{a['severity']}] {a['month']}: {a['detail']}")
        print(f"\n  Gemini: {summary.gemini_verdict}")
        print(f"{'═'*60}\n")


# ══════════════════════════════════════════════════════════════
# USAGE EXAMPLES
# ══════════════════════════════════════════════════════════════

def example_basic():
    """ตัวอย่างการใช้งานพื้นฐาน"""
    engine = FinancialEngine(
        gemini_api_key  = "YOUR_API_KEY",   # หรือ userdata.get('GEMINI_API_KEY')
        target_year     = "68",
        retained_begin  = 2278127.06,        # กำไรสะสมต้นปี
        admin_expense   = 5000.,             # ค่าใช้จ่ายบริหาร
        inventory_begin = 409280.74,         # สินค้าต้นงวด
        inventory_end   = 409280.74,         # สินค้าปลายงวด
    )
    summary = engine.process("ภพ.30_บริษัทXYZ.xlsx")
    engine.print_summary(summary)
    # engine.export(summary, "output.xlsx")
    return summary


def example_colab():
    """ตัวอย่างสำหรับ Google Colab"""
    from google.colab import files, userdata
    print("📂 อัปโหลดไฟล์ ภพ.30...")
    uploaded = files.upload()
    filepath = list(uploaded.keys())[0]

    engine = FinancialEngine(
        gemini_api_key = userdata.get('GEMINI_API_KEY'),
        target_year    = "68",
        retained_begin = 0.,     # ← ปรับตามลูกค้า
        admin_expense  = 0.,     # ← ปรับตามลูกค้า
    )
    summary = engine.process(filepath)
    engine.print_summary(summary)

    if summary.gemini_verdict in ('ready', 'skip'):
        print("✅ พร้อม Export")
        # engine.export(summary)
    else:
        print("⚠️  กรุณาตรวจสอบก่อน Export")
    return summary


if __name__ == "__main__":
    import sys
    # กัน Colab inject '-f' หรือ path kernel เข้ามาเป็น argv
    args = [a for a in sys.argv[1:]
            if not a.startswith('-') and (a.endswith('.xlsx') or a.endswith('.xls'))]
    if args:
        engine = FinancialEngine(
            gemini_api_key = "dummy",
            target_year    = "68",
        )
        summary = engine.process(args[0], skip_gemini=True)
        engine.print_summary(summary)
    else:
        print("Usage: python financial_engine.py <ภพ30_file.xlsx>")
