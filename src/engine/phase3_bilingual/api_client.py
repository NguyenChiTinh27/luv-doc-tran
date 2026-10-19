"""
src/engine/phase3_bilingual/api_client.py
──────────────────────────────────────────
Optional external LLM client adapter for DeepSeek / OpenAI translation (F5 / R2).

Active only when user explicitly supplies credentials via configuration or CLI flags
(--online, --api-key). In default offline mode, zero external API calls are made.
"""

import os
import time
import json
import logging
from typing import Optional, List, Dict, Any
from enum import Enum

from src.engine.phase3_bilingual.rule_synthesizer import RuleBasedSynthesizer
from src.engine.phase3_bilingual.cost_guard import ApiCostViolationError

logger = logging.getLogger("luv_doctran.api_client")


class LLMProvider(str, Enum):
    DEEPSEEK = "deepseek"
    OPENAI = "openai"


class ExternalLLMClient:
    """
    Adapter client for external translation LLMs (DeepSeek, OpenAI).
    Only enabled when an explicit API key is provided.
    """

    DEFAULT_BASE_URLS: Dict[str, str] = {
        LLMProvider.DEEPSEEK.value: "https://api.deepseek.com/v1",
        LLMProvider.OPENAI.value: "https://api.openai.com/v1",
    }

    DEFAULT_MODELS: Dict[str, str] = {
        LLMProvider.DEEPSEEK.value: "deepseek-chat",
        LLMProvider.OPENAI.value: "gpt-4o-mini",
    }

    SYSTEM_PROMPT: str = (
        "You are an expert technical translator specializing in chemical safety data sheets "
        "(MSDS/SDS), technical data sheets (TDS), and legal contracts. Translate the provided "
        "English source text into accurate, professional Vietnamese. Maintain all CAS numbers, "
        "chemical formulas, percentages, units, and numeric values exactly without modification. "
        "Do not add commentary; output only the direct Vietnamese translation."
    )

    def __init__(
        self,
        api_key: Optional[str] = None,
        provider: str = LLMProvider.DEEPSEEK.value,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        timeout: float = 30.0,
        max_retries: int = 3,
        fallback_to_offline: bool = True
    ):
        self.provider = provider.lower() if provider else LLMProvider.DEEPSEEK.value
        resolved_key = api_key or os.environ.get("DEEPSEEK_API_KEY" if self.provider == "deepseek" else "OPENAI_API_KEY")
        if not resolved_key and self.provider == "deepseek":
            root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
            settings_path = os.path.join(root_dir, "settings.json")
            if os.path.exists(settings_path):
                try:
                    with open(settings_path, "r", encoding="utf-8") as sf:
                        s_data = json.load(sf)
                        resolved_key = s_data.get("deepseek_api_key")
                except Exception:
                    pass
            if not resolved_key:
                env_path = os.path.join(root_dir, ".env")
                if os.path.exists(env_path):
                    try:
                        with open(env_path, "r", encoding="utf-8") as ef:
                            for eline in ef:
                                if eline.startswith("DEEPSEEK_API_KEY="):
                                    resolved_key = eline.split("=", 1)[1].strip()
                                    break
                    except Exception:
                        pass
        self.api_key = resolved_key
        self.base_url = (base_url or self.DEFAULT_BASE_URLS.get(self.provider, "https://api.deepseek.com/v1")).rstrip("/")
        self.model = model or self.DEFAULT_MODELS.get(self.provider, "deepseek-chat")
        self.timeout = timeout
        self.max_retries = max_retries
        self.fallback_to_offline = fallback_to_offline

    def is_configured(self) -> bool:
        """Returns True if an API key is available."""
        return bool(self.api_key and self.api_key.strip())

    def translate(self, text: str, context: Optional[str] = None) -> str:
        """
        Translates a single English text string into Vietnamese.
        If not configured, falls back to RuleBasedSynthesizer if enabled,
        or raises ValueError.
        """
        if not text or not text.strip():
            return ""

        if not self.is_configured():
            if self.fallback_to_offline:
                return RuleBasedSynthesizer.synthesize(text)
            raise ValueError(f"ExternalLLMClient for {self.provider} requires an API key.")

        # Attempt API translation with retries
        import requests

        endpoint = f"{self.base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "User-Agent": "LUV-DOC-TRAN/1.1.0"
        }

        user_content = text
        if context:
            user_content = f"Context: {context}\n\nSource text:\n{text}"

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": self.SYSTEM_PROMPT},
                {"role": "user", "content": user_content}
            ],
            "temperature": 0.1,
            "max_tokens": 1024
        }

        last_error = None
        for attempt in range(self.max_retries):
            try:
                resp = requests.post(
                    endpoint,
                    headers=headers,
                    json=payload,
                    timeout=self.timeout
                )
                if resp.status_code == 200:
                    data = resp.json()
                    translated = data["choices"][0]["message"]["content"].strip()
                    return translated
                else:
                    logger.warning(
                        "LLM API request failed (%s): %s",
                        resp.status_code,
                        resp.text[:200]
                    )
            except ApiCostViolationError:
                # Re-raise cost guard violation immediately
                raise
            except Exception as e:
                last_error = e
                logger.warning("LLM API attempt %d error: %s", attempt + 1, e)

            if attempt < self.max_retries - 1:
                time.sleep(1.0 * (attempt + 1))

        if self.fallback_to_offline:
            logger.info("Falling back to RuleBasedSynthesizer due to API error: %s", last_error)
            return RuleBasedSynthesizer.synthesize(text)

        raise RuntimeError(f"External LLM translation failed after {self.max_retries} attempts: {last_error}")

    def batch_translate(self, texts: List[str], context: Optional[str] = None) -> List[str]:
        """Translates a batch of strings."""
        return [self.translate(t, context=context) for t in texts]
