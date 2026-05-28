"""
ai_provider.py
═══════════════════════════════════════════════════════════════
AI Provider Abstraction Layer
ออกแบบให้ swap AI ได้โดยไม่แก้ Engine เลย

เปลี่ยน Model: แค่แก้ config.yml บรรทัดเดียว
  provider: "gemini"   # หรือ "openai" / "claude" / "local"

โครงสร้าง:
  AIProvider          — Abstract base class (interface)
  GeminiProvider      — Google Gemini
  OpenAIProvider      — OpenAI GPT
  ClaudeProvider      — Anthropic Claude
  LocalProvider       — Ollama / vLLM (fine-tuned model)
  AIProviderFactory   — สร้าง provider จาก config
  TrainingDataCollector — เก็บ feedback สำหรับ fine-tuning
═══════════════════════════════════════════════════════════════
"""

import os
import json
import time
import logging
import hashlib
from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional

log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════
# DATA STRUCTURES
# ══════════════════════════════════════════════════════════════

@dataclass
class AIRequest:
    """Request ที่ส่งให้ AI"""
    task: str                    # 'analyze'|'classify'|'extract'|'summarize'
    prompt: str                  # prompt หลัก
    context: dict = field(default_factory=dict)   # ข้อมูล context เพิ่มเติม
    max_tokens: int = 2000
    temperature: float = 0.1     # ต่ำ = conservative (เหมาะกับบัญชี)
    request_id: str = ""

    def __post_init__(self):
        if not self.request_id:
            self.request_id = hashlib.md5(
                f"{self.task}{self.prompt[:50]}{time.time()}".encode()
            ).hexdigest()[:8]


@dataclass
class AIResponse:
    """Response จาก AI"""
    text: str                    # ข้อความตอบ
    provider: str                # ชื่อ provider ที่ใช้
    model: str                   # ชื่อ model ที่ใช้
    request_id: str = ""
    tokens_used: int = 0
    latency_ms: float = 0.
    confidence: float = 1.0      # ความมั่นใจ (ถ้า provider รองรับ)
    raw_response: dict = field(default_factory=dict)
    error: str = ""

    @property
    def success(self) -> bool:
        return not self.error and bool(self.text)


@dataclass
class TrainingExample:
    """ตัวอย่างสำหรับ fine-tuning"""
    task: str
    input_prompt: str
    expected_output: str         # คำตอบที่ถูกต้อง (จากนักบัญชี)
    actual_output: str           # คำตอบที่ AI ให้มา
    is_correct: bool             # นักบัญชียืนยันว่าถูก
    correction: str = ""         # ถ้าผิด นักบัญชีแก้ยังไง
    company_type: str = ""       # ประเภทธุรกิจ
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    metadata: dict = field(default_factory=dict)


# ══════════════════════════════════════════════════════════════
# SYSTEM PROMPT (ใช้ร่วมกันทุก Provider)
# ══════════════════════════════════════════════════════════════

ACCOUNTING_SYSTEM_PROMPT = """
คุณคือผู้ช่วยนักบัญชีผู้เชี่ยวชาญด้านบัญชีและภาษีสำหรับธุรกิจ SME ไทย
มีความเชี่ยวชาญใน:
- ภาษีมูลค่าเพิ่ม (VAT) และ ภ.พ.30
- ภาษีเงินได้นิติบุคคล (CIT) ภงด.50/51
- มาตรฐานการบัญชีไทย (NPAEs/PAEs)
- การตรวจสอบงบการเงินและจุดเสี่ยง

หลักการตอบ:
1. ตอบเป็นภาษาไทย กระชับ ตรงประเด็น
2. ระบุตัวเลขสำคัญเสมอ
3. ถ้าข้อมูลไม่เพียงพอ บอกตรงๆ อย่าเดา
4. ใส่ VERDICT: ท้ายสุดเสมอเมื่อถามเรื่องการตรวจสอบ
"""


# ══════════════════════════════════════════════════════════════
# 1. Abstract Base Class — AIProvider
# ══════════════════════════════════════════════════════════════

class AIProvider(ABC):
    """
    Interface หลักที่ทุก Provider ต้อง implement
    Engine จะเรียกผ่าน interface นี้เท่านั้น
    → swap provider โดยไม่แก้ Engine เลย
    """

    def __init__(self, config: dict):
        self.config     = config
        self.model_name = config.get('model', 'unknown')
        self.timeout    = config.get('timeout', 60)
        self.max_retries = config.get('max_retries', 3)
        self._call_count = 0
        self._error_count = 0

    @abstractmethod
    def _call_api(self, request: AIRequest) -> AIResponse:
        """เรียก API จริงๆ — ต้อง implement ใน subclass"""
        pass

    @abstractmethod
    def health_check(self) -> bool:
        """ตรวจสอบว่า provider พร้อมใช้งาน"""
        pass

    def call(self, request: AIRequest) -> AIResponse:
        """
        เรียก AI พร้อม retry + logging + timing
        ทุก Provider ใช้ method นี้ผ่าน base class
        """
        start = time.time()
        self._call_count += 1

        for attempt in range(self.max_retries):
            try:
                response = self._call_api(request)
                response.latency_ms = (time.time() - start) * 1000
                response.request_id = request.request_id

                log.info(f"[{self.__class__.__name__}] task={request.task} "
                         f"latency={response.latency_ms:.0f}ms "
                         f"tokens={response.tokens_used}")
                return response

            except Exception as e:
                err_str = str(e)
                self._error_count += 1

                # Rate limit → รอแล้ว retry
                if any(code in err_str for code in ['429', 'rate', 'quota']):
                    wait = 15 * (attempt + 1)
                    log.warning(f"Rate limit — รอ {wait}s (attempt {attempt+1})")
                    time.sleep(wait)
                    continue

                # Error อื่นๆ
                log.error(f"[{self.__class__.__name__}] Error: {e}")
                if attempt == self.max_retries - 1:
                    return AIResponse(
                        text="", provider=self.__class__.__name__,
                        model=self.model_name, error=str(e)
                    )
                time.sleep(2)

        return AIResponse(
            text="", provider=self.__class__.__name__,
            model=self.model_name, error="Max retries exceeded"
        )

    # ── Convenience methods (เรียกจาก Engine) ──

    def analyze(self, data_text: str, company_info: str = "") -> AIResponse:
        """วิเคราะห์งบการเงิน"""
        req = AIRequest(
            task="analyze",
            prompt=f"{ACCOUNTING_SYSTEM_PROMPT}\n\n{company_info}\n\n{data_text}"
        )
        return self.call(req)

    def classify_account(self, account_name: str,
                          amount: float, context: str = "") -> AIResponse:
        """จัดหมวดบัญชี"""
        prompt = f"""
{ACCOUNTING_SYSTEM_PROMPT}

จัดหมวดบัญชีต่อไปนี้:
ชื่อบัญชี: {account_name}
จำนวน: {amount:,.2f} บาท
context: {context}

ตอบในรูป JSON:
{{
  "account_type": "สินทรัพย์|หนี้สิน|ทุน|รายได้|ค่าใช้จ่าย",
  "sub_type": "...",
  "is_vat_applicable": true/false,
  "confidence": 0.0-1.0,
  "note": "..."
}}
"""
        return self.call(AIRequest(task="classify", prompt=prompt,
                                   temperature=0.0))

    def extract_summary(self, raw_text: str) -> AIResponse:
        """สกัดตัวเลขสรุปจาก text ดิบ"""
        prompt = f"""
{ACCOUNTING_SYSTEM_PROMPT}

สกัดตัวเลขจาก text ต่อไปนี้:
{raw_text}

ตอบใน JSON:
{{
  "sale_vat": 0.0,
  "sale_exempt": 0.0,
  "purchase": 0.0,
  "vat_sale": 0.0,
  "vat_buy": 0.0,
  "vat_net": 0.0
}}
ถ้าไม่มีข้อมูลใส่ 0.0 ห้ามเดา
"""
        return self.call(AIRequest(task="extract", prompt=prompt,
                                   temperature=0.0))

    def get_stats(self) -> dict:
        return {
            'provider': self.__class__.__name__,
            'model': self.model_name,
            'total_calls': self._call_count,
            'errors': self._error_count,
            'error_rate': self._error_count / max(self._call_count, 1),
        }


# ══════════════════════════════════════════════════════════════
# 2. GeminiProvider
# ══════════════════════════════════════════════════════════════

class GeminiProvider(AIProvider):
    """Google Gemini — ใช้งานอยู่ตอนนี้"""

    def __init__(self, config: dict):
        super().__init__(config)
        try:
            import google.generativeai as genai
            genai.configure(api_key=config['api_key'])
            self._model = genai.GenerativeModel(
                self.model_name,
                generation_config={
                    'temperature': config.get('temperature', 0.1),
                    'max_output_tokens': config.get('max_tokens', 2000),
                }
            )
            log.info(f"GeminiProvider ready: {self.model_name}")
        except ImportError:
            raise ImportError("pip install google-generativeai")

    def _call_api(self, request: AIRequest) -> AIResponse:
        resp = self._model.generate_content(request.prompt)
        return AIResponse(
            text=resp.text,
            provider="gemini",
            model=self.model_name,
            tokens_used=getattr(resp, 'usage_metadata', {}).get(
                'total_token_count', 0) if hasattr(resp, 'usage_metadata') else 0,
        )

    def health_check(self) -> bool:
        try:
            r = self._model.generate_content("ping")
            return bool(r.text)
        except:
            return False


# ══════════════════════════════════════════════════════════════
# 3. OpenAIProvider
# ══════════════════════════════════════════════════════════════

class OpenAIProvider(AIProvider):
    """OpenAI GPT-4o / GPT-4 — swap ได้ทันที"""

    def __init__(self, config: dict):
        super().__init__(config)
        try:
            from openai import OpenAI
            self._client = OpenAI(api_key=config['api_key'])
            log.info(f"OpenAIProvider ready: {self.model_name}")
        except ImportError:
            raise ImportError("pip install openai")

    def _call_api(self, request: AIRequest) -> AIResponse:
        resp = self._client.chat.completions.create(
            model=self.model_name,
            messages=[
                {"role": "system", "content": ACCOUNTING_SYSTEM_PROMPT},
                {"role": "user",   "content": request.prompt}
            ],
            temperature=request.temperature,
            max_tokens=request.max_tokens,
        )
        return AIResponse(
            text=resp.choices[0].message.content,
            provider="openai",
            model=self.model_name,
            tokens_used=resp.usage.total_tokens,
        )

    def health_check(self) -> bool:
        try:
            r = self._client.chat.completions.create(
                model=self.model_name,
                messages=[{"role": "user", "content": "ping"}],
                max_tokens=5
            )
            return bool(r.choices[0].message.content)
        except:
            return False


# ══════════════════════════════════════════════════════════════
# 4. ClaudeProvider
# ══════════════════════════════════════════════════════════════

class ClaudeProvider(AIProvider):
    """Anthropic Claude — swap ได้ทันที"""

    def __init__(self, config: dict):
        super().__init__(config)
        try:
            import anthropic
            self._client = anthropic.Anthropic(api_key=config['api_key'])
            log.info(f"ClaudeProvider ready: {self.model_name}")
        except ImportError:
            raise ImportError("pip install anthropic")

    def _call_api(self, request: AIRequest) -> AIResponse:
        resp = self._client.messages.create(
            model=self.model_name,
            max_tokens=request.max_tokens,
            system=ACCOUNTING_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": request.prompt}]
        )
        return AIResponse(
            text=resp.content[0].text,
            provider="claude",
            model=self.model_name,
            tokens_used=resp.usage.input_tokens + resp.usage.output_tokens,
        )

    def health_check(self) -> bool:
        try:
            r = self._client.messages.create(
                model=self.model_name,
                max_tokens=5,
                messages=[{"role": "user", "content": "ping"}]
            )
            return bool(r.content[0].text)
        except:
            return False


# ══════════════════════════════════════════════════════════════
# 5. LocalProvider — Ollama / fine-tuned model
# ══════════════════════════════════════════════════════════════

class LocalProvider(AIProvider):
    """
    รัน model บน local server (Ollama / vLLM / LM Studio)
    ใช้สำหรับ fine-tuned model ที่เทรนเองจากข้อมูลบัญชีไทย
    """

    def __init__(self, config: dict):
        super().__init__(config)
        self.base_url = config.get('base_url', 'http://localhost:11434')
        log.info(f"LocalProvider ready: {self.base_url}/{self.model_name}")

    def _call_api(self, request: AIRequest) -> AIResponse:
        import urllib.request
        payload = json.dumps({
            "model":  self.model_name,
            "prompt": request.prompt,
            "stream": False,
            "options": {"temperature": request.temperature}
        }).encode()

        req = urllib.request.Request(
            f"{self.base_url}/api/generate",
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=self.timeout) as r:
            data = json.loads(r.read())

        return AIResponse(
            text=data.get("response", ""),
            provider="local",
            model=self.model_name,
            tokens_used=data.get("eval_count", 0),
        )

    def health_check(self) -> bool:
        try:
            import urllib.request
            with urllib.request.urlopen(
                f"{self.base_url}/api/tags", timeout=5
            ) as r:
                return r.status == 200
        except:
            return False


# ══════════════════════════════════════════════════════════════
# 6. AIProviderFactory — สร้าง provider จาก config
# ══════════════════════════════════════════════════════════════

class AIProviderFactory:
    """
    สร้าง Provider จาก config
    Engine เรียก Factory.create() ครั้งเดียวตอนเริ่มต้น

    config.yml ตัวอย่าง:
    ─────────────────────
    ai:
      provider: gemini          # เปลี่ยนบรรทัดนี้บรรทัดเดียว
      fallback: openai          # ถ้า primary ล้ม ใช้ fallback
      gemini:
        api_key: "AIza..."
        model: "gemini-2.5-flash"
        max_retries: 3
      openai:
        api_key: "sk-..."
        model: "gpt-4o"
      claude:
        api_key: "sk-ant-..."
        model: "claude-opus-4-5"
      local:
        base_url: "http://localhost:11434"
        model: "thai-accounting-v1"   # ชื่อ fine-tuned model
    """

    PROVIDERS = {
        'gemini': GeminiProvider,
        'openai': OpenAIProvider,
        'claude': ClaudeProvider,
        'local':  LocalProvider,
    }

    @classmethod
    def create(cls, config: dict) -> AIProvider:
        """สร้าง primary provider จาก config"""
        provider_name = config.get('provider', 'gemini').lower()
        provider_config = config.get(provider_name, {})

        # inject api_key จาก env ถ้าไม่มีใน config
        if 'api_key' not in provider_config:
            env_key = f"{provider_name.upper()}_API_KEY"
            provider_config['api_key'] = os.environ.get(env_key, '')

        provider_class = cls.PROVIDERS.get(provider_name)
        if not provider_class:
            raise ValueError(f"ไม่รู้จัก provider: {provider_name} "
                             f"(รองรับ: {list(cls.PROVIDERS.keys())})")

        log.info(f"สร้าง AI Provider: {provider_name}")
        return provider_class(provider_config)

    @classmethod
    def create_with_fallback(cls, config: dict) -> 'AIProviderWithFallback':
        """สร้าง provider พร้อม fallback อัตโนมัติ"""
        primary  = cls.create(config)
        fallback = None

        fallback_name = config.get('fallback')
        if fallback_name and fallback_name != config.get('provider'):
            try:
                fallback_config = dict(config)
                fallback_config['provider'] = fallback_name
                fallback = cls.create(fallback_config)
                log.info(f"Fallback provider: {fallback_name}")
            except Exception as e:
                log.warning(f"ไม่สามารถสร้าง fallback provider: {e}")

        return AIProviderWithFallback(primary, fallback)

    @classmethod
    def from_yaml(cls, yaml_path: str) -> 'AIProviderWithFallback':
        """โหลด config จาก YAML file"""
        try:
            import yaml
            with open(yaml_path) as f:
                config = yaml.safe_load(f)
            return cls.create_with_fallback(config.get('ai', {}))
        except ImportError:
            raise ImportError("pip install pyyaml")

    @classmethod
    def from_colab_secrets(cls, provider: str = 'gemini',
                           model: str = 'gemini-2.5-flash') -> AIProvider:
        """สร้างจาก Google Colab Secrets — ง่ายที่สุดสำหรับ Colab"""
        try:
            from google.colab import userdata
            api_key = userdata.get(f'{provider.upper()}_API_KEY')
        except Exception:
            api_key = os.environ.get(f'{provider.upper()}_API_KEY', '')

        config = {
            'provider': provider,
            provider: {'api_key': api_key, 'model': model}
        }
        return cls.create(config)


# ══════════════════════════════════════════════════════════════
# 7. AIProviderWithFallback — Auto failover
# ══════════════════════════════════════════════════════════════

class AIProviderWithFallback(AIProvider):
    """
    Wrapper ที่ fallback ไป provider สำรองอัตโนมัติ
    ถ้า primary ล้ม (error / rate limit)
    """

    def __init__(self, primary: AIProvider, fallback: Optional[AIProvider] = None):
        self.primary  = primary
        self.fallback = fallback
        self.model_name = primary.model_name
        self.config = {}
        self._call_count = 0
        self._error_count = 0

    def _call_api(self, request: AIRequest) -> AIResponse:
        # ไม่ถูกเรียกโดยตรง — override call() แทน
        return self.primary._call_api(request)

    def call(self, request: AIRequest) -> AIResponse:
        self._call_count += 1
        resp = self.primary.call(request)

        if not resp.success and self.fallback:
            log.warning(f"Primary failed ({resp.error}) → ใช้ fallback")
            resp = self.fallback.call(request)

        return resp

    def health_check(self) -> bool:
        ok = self.primary.health_check()
        if not ok and self.fallback:
            log.warning("Primary unhealthy → ตรวจ fallback")
            return self.fallback.health_check()
        return ok

    # delegate convenience methods
    def analyze(self, *args, **kwargs):
        return super().analyze(*args, **kwargs)


# ══════════════════════════════════════════════════════════════
# 8. TrainingDataCollector — เก็บ feedback สำหรับ fine-tuning
# ══════════════════════════════════════════════════════════════

class TrainingDataCollector:
    """
    เก็บ feedback จากนักบัญชีเพื่อ fine-tune model ในอนาคต

    วิธีใช้:
        collector = TrainingDataCollector("training_data/")
        # หลัง AI วิเคราะห์
        collector.record(
            task="analyze",
            input_prompt=prompt,
            ai_output=response.text,
            is_correct=True,          # นักบัญชีกด approve
            correction="",            # ถ้า False ใส่การแก้ไข
            company_type="ค้าปลีก"
        )
        # export เพื่อ fine-tune
        collector.export_jsonl("fine_tune_data.jsonl")
    """

    def __init__(self, data_dir: str = "training_data"):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self._examples: list = []
        self._load_existing()

    def _load_existing(self):
        """โหลดข้อมูลที่มีอยู่แล้ว"""
        data_file = self.data_dir / "examples.jsonl"
        if data_file.exists():
            with open(data_file) as f:
                for line in f:
                    try:
                        self._examples.append(json.loads(line))
                    except:
                        pass
            log.info(f"โหลด training examples: {len(self._examples)} รายการ")

    def record(self, task: str, input_prompt: str,
               ai_output: str, is_correct: bool,
               correction: str = "", company_type: str = "",
               metadata: dict = None) -> TrainingExample:
        """บันทึก 1 ตัวอย่าง"""
        example = TrainingExample(
            task=task,
            input_prompt=input_prompt,
            expected_output=correction if not is_correct else ai_output,
            actual_output=ai_output,
            is_correct=is_correct,
            correction=correction,
            company_type=company_type,
            metadata=metadata or {}
        )
        self._examples.append(asdict(example))
        self._save(asdict(example))

        log.info(f"บันทึก example: task={task} correct={is_correct}")
        return example

    def _save(self, example: dict):
        """append ลง jsonl file"""
        with open(self.data_dir / "examples.jsonl", 'a', encoding='utf-8') as f:
            f.write(json.dumps(example, ensure_ascii=False) + "\n")

    def export_jsonl(self, output_path: str,
                     only_correct: bool = False,
                     task_filter: str = None) -> int:
        """
        Export เป็น JSONL format สำหรับ fine-tuning
        รองรับ format ของ OpenAI / Gemini fine-tuning
        """
        examples = self._examples
        if only_correct:
            examples = [e for e in examples if e['is_correct']]
        if task_filter:
            examples = [e for e in examples if e['task'] == task_filter]

        with open(output_path, 'w', encoding='utf-8') as f:
            for ex in examples:
                # Format: OpenAI fine-tuning format
                fine_tune_item = {
                    "messages": [
                        {"role": "system",    "content": ACCOUNTING_SYSTEM_PROMPT},
                        {"role": "user",      "content": ex['input_prompt']},
                        {"role": "assistant", "content": ex['expected_output']},
                    ],
                    "metadata": {
                        "task": ex['task'],
                        "company_type": ex['company_type'],
                        "created_at": ex['created_at'],
                    }
                }
                f.write(json.dumps(fine_tune_item, ensure_ascii=False) + "\n")

        log.info(f"Export {len(examples)} examples → {output_path}")
        return len(examples)

    def get_stats(self) -> dict:
        """สรุปสถิติ training data"""
        total    = len(self._examples)
        correct  = sum(1 for e in self._examples if e['is_correct'])
        by_task  = {}
        by_type  = {}
        for e in self._examples:
            by_task[e['task']] = by_task.get(e['task'], 0) + 1
            by_type[e['company_type']] = by_type.get(e['company_type'], 0) + 1
        return {
            'total': total,
            'correct': correct,
            'accuracy': correct / max(total, 1),
            'by_task': by_task,
            'by_company_type': by_type,
            'ready_for_finetune': total >= 50,  # ต้องการ ≥50 examples
        }

    def print_stats(self):
        s = self.get_stats()
        print(f"\n{'─'*45}")
        print(f"  Training Data Stats")
        print(f"{'─'*45}")
        print(f"  รวมทั้งหมด:  {s['total']} examples")
        print(f"  ถูกต้อง:     {s['correct']} ({s['accuracy']:.0%})")
        print(f"  by task:     {s['by_task']}")
        print(f"  by type:     {s['by_company_type']}")
        print(f"  พร้อม fine-tune: {'✅' if s['ready_for_finetune'] else '❌ (ต้องการ ≥50)'}")
        print(f"{'─'*45}\n")


# ══════════════════════════════════════════════════════════════
# CONFIG TEMPLATE — config.yml
# ══════════════════════════════════════════════════════════════

CONFIG_TEMPLATE = """
# config.yml — Financial Engine AI Configuration
# เปลี่ยน provider แค่บรรทัดเดียว ไม่ต้องแก้ code เลย

ai:
  provider: gemini        # gemini | openai | claude | local
  fallback: openai        # ใช้เมื่อ primary ล้ม (optional)

  gemini:
    model: gemini-2.5-flash
    max_retries: 3
    timeout: 60
    # api_key: ใส่ใน GEMINI_API_KEY env var หรือ Colab Secrets

  openai:
    model: gpt-4o
    max_retries: 3
    timeout: 60
    # api_key: ใส่ใน OPENAI_API_KEY env var

  claude:
    model: claude-opus-4-5
    max_retries: 3
    timeout: 60
    # api_key: ใส่ใน CLAUDE_API_KEY env var

  local:
    base_url: http://localhost:11434
    model: thai-accounting-v1    # ชื่อ fine-tuned model ใน Ollama
    timeout: 120

engine:
  target_year: "68"
  business_type: ทั่วไป          # ทั่วไป | รถยนต์ | ค้าปลีก | บริการ

training:
  data_dir: training_data/
  min_examples_for_finetune: 50
  auto_collect: true             # เก็บ feedback อัตโนมัติ
"""

def write_config_template(path: str = "config.yml"):
    """เขียน config.yml template"""
    with open(path, 'w', encoding='utf-8') as f:
        f.write(CONFIG_TEMPLATE)
    print(f"✅ เขียน config template: {path}")


# ══════════════════════════════════════════════════════════════
# INTEGRATION กับ FinancialEngine (ที่มีอยู่แล้ว)
# ══════════════════════════════════════════════════════════════

class GeminiAnalyzerV2:
    """
    Drop-in replacement สำหรับ GeminiAnalyzer เดิม
    ใช้ AIProvider แทน genai โดยตรง
    → financial_engine.py ไม่ต้องแก้เลย ยกเว้นเปลี่ยน import
    """

    def __init__(self, provider: AIProvider,
                 collector: Optional[TrainingDataCollector] = None):
        self.provider  = provider
        self.collector = collector

    def analyze_vat(self, company, monthly, classification,
                    anomalies, calc_result) -> tuple:
        """วิเคราะห์ข้อมูล ภพ.30 — interface เดิม ทำงานเหมือนเดิม"""
        valid = [m for m in monthly if m.total_sale > 0]
        pl    = calc_result['pl']
        vat   = classification['tax_analysis']

        month_lines = "\n".join([
            f"- {m.month_full}: ขาย {m.sale_vat:,.0f} ยกเว้น {m.sale_exempt:,.0f}"
            f" | ซื้อ {m.purchase:,.0f} | ภาษี {m.vat_net:,.2f} [conf={m.confidence:.0%}]"
            for m in valid
        ])
        anomaly_text = "\n".join([
            f"  ⚠ [{a['severity']}] {a['month']}: {a['detail']}"
            for a in anomalies
        ]) or "  (ไม่พบความผิดปกติ)"

        prompt = f"""
วิเคราะห์ข้อมูล ภพ.30:

บริษัท: {company.name} | เลขผู้เสียภาษี: {company.tax_id}
ปีภาษี: 25{company.fiscal_year}

ข้อมูลรายเดือน:
{month_lines}

รวมทั้งปี: รายได้ {pl['total_revenue']:,.2f} | กำไรสุทธิ {pl['net_profit']:,.2f}
VAT สุทธิ: {vat['total_vat_net']:,.2f}

ความผิดปกติ:
{anomaly_text}

วิเคราะห์ 5 ด้าน:
1. ความสมเหตุสมผลรายเดือน
2. ความถูกต้องของภาษี VAT
3. อัตรากำไรและต้นทุน
4. จุดเสี่ยงสำคัญ
5. สรุปและคำแนะนำ

ท้ายสุดระบุ:
VERDICT: พร้อมปิดงบ
หรือ
VERDICT: ต้องตรวจสอบเพิ่มเติม — [ระบุ]
"""
        resp = self.provider.call(AIRequest(task="analyze", prompt=prompt))

        # บันทึกลง training data ถ้ามี collector
        if self.collector:
            self.collector.record(
                task="analyze",
                input_prompt=prompt,
                ai_output=resp.text,
                is_correct=True,  # default true — นักบัญชีแก้ทีหลังได้
                company_type=getattr(company, 'business_type', '')
            )

        verdict = ('ready' if 'VERDICT: พร้อมปิดงบ' in resp.text
                   else 'review' if 'VERDICT:' in resp.text else 'error')
        return resp.text, verdict


# ══════════════════════════════════════════════════════════════
# QUICK START — ใช้งานใน Colab
# ══════════════════════════════════════════════════════════════

def quick_start_colab(provider_name: str = 'gemini',
                      model: str = 'gemini-2.5-flash') -> AIProvider:
    """
    เริ่มต้นใช้งานใน Colab — 1 บรรทัด

    ตัวอย่าง:
        from ai_provider import quick_start_colab
        ai = quick_start_colab('gemini')           # ใช้ Gemini
        ai = quick_start_colab('openai', 'gpt-4o') # เปลี่ยนเป็น OpenAI
        ai = quick_start_colab('local', 'my-model') # ใช้ local model
    """
    provider = AIProviderFactory.from_colab_secrets(provider_name, model)
    ok = provider.health_check()
    status = "✅ พร้อม" if ok else "⚠️  ไม่สามารถเชื่อมต่อ"
    print(f"{status} | Provider: {provider_name} | Model: {model}")
    return provider


if __name__ == "__main__":
    # ทดสอบ — แสดง config template
    print("=" * 55)
    print("  AI Provider Abstraction Layer")
    print("  เปลี่ยน AI ได้โดยไม่แก้ Engine")
    print("=" * 55)
    print("\nProviders ที่รองรับ:")
    for name in AIProviderFactory.PROVIDERS:
        print(f"  - {name}")
    print("\nสร้าง config.yml:")
    write_config_template("config.yml")
    print("\nวิธีใช้ใน Colab:")
    print("  from ai_provider import quick_start_colab")
    print("  ai = quick_start_colab('gemini')")
    print("  ai = quick_start_colab('openai', 'gpt-4o')")
