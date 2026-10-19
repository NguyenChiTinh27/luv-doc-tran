"""
src/engine/phase3_bilingual/__init__.py
───────────────────────────────────────
Phase 3: Bilingual Engine & Pairing Typography Package.

Exports:
  - BilingualTranslator: Main orchestrating translator
  - PairingEngine: Pattern A & Pattern B typography builder
  - PairingPattern: Enum (PATTERN_A, PATTERN_B)
  - ZeroCostGuard: Runtime guard blocking external paid API calls
  - ApiCostViolationError: Exception raised on paid API attempt
  - OfflineGlossary: 16 GHS headers, 60+ chemicals, mandatory HSE phrases
  - TokenSanitizer: Pre-translation token spacing & OCR typo cleaner
  - LeakHunter: Post-translation English ordinal & heading repairer
  - RuleBasedSynthesizer: 100% offline zero-cost domain synthesizer
  - ExternalLLMClient: Optional adapter for DeepSeek / OpenAI LLMs
  - Typography constants: DARK_NAVY_HEX, DARK_NAVY_RGB, etc.
"""

from src.engine.phase3_bilingual.cost_guard import (
    ZeroCostGuard,
    ApiCostViolationError
)
from src.engine.phase3_bilingual.typography import (
    DARK_NAVY_HEX,
    DARK_NAVY_RGB,
    DARK_BLUE_HEX,
    DARK_BLUE_RGB,
    BLACK_HEX,
    BLACK_RGB,
    DEFAULT_FONT_NAME,
    DEFAULT_BODY_FONT_SIZE_PT,
    DEFAULT_TABLE_FONT_SIZE_PT,
    SPACING_AFTER_DXA,
    SPACING_AFTER_PT,
    ROMAN_NUMERALS,
    add_keep_next,
    set_paragraph_spacing,
    set_cell_valign_center,
    clean_inherited_styles
)
from src.engine.phase3_bilingual.offline_glossary import OfflineGlossary
from src.engine.phase3_bilingual.token_sanitizer import TokenSanitizer
from src.engine.phase3_bilingual.leak_hunter import LeakHunter
from src.engine.phase3_bilingual.rule_synthesizer import RuleBasedSynthesizer
from src.engine.phase3_bilingual.pairing_engine import (
    PairingEngine,
    PairingPattern,
    BodyPair,
    CellPair,
    PairedDocument,
    PairingBuilder  # backward-compat alias
)
from src.engine.phase3_bilingual.api_client import (
    ExternalLLMClient,
    LLMProvider
)
from src.engine.phase3_bilingual.bilingual_translator import (
    BilingualTranslator,
    TranslationConfig
)

__all__ = [
    "BilingualTranslator",
    "PairingEngine",
    "PairingPattern",
    "ZeroCostGuard",
    "OfflineGlossary",
    "TokenSanitizer",
    "LeakHunter",
    "RuleBasedSynthesizer",
    # Additional components
    "ApiCostViolationError",
    "ExternalLLMClient",
    "LLMProvider",
    "TranslationConfig",
    "BodyPair",
    "CellPair",
    "PairedDocument",
    "PairingBuilder",
    "DARK_NAVY_HEX",
    "DARK_NAVY_RGB",
    "DARK_BLUE_HEX",
    "DARK_BLUE_RGB",
    "BLACK_HEX",
    "BLACK_RGB",
    "DEFAULT_FONT_NAME",
    "DEFAULT_BODY_FONT_SIZE_PT",
    "DEFAULT_TABLE_FONT_SIZE_PT",
    "SPACING_AFTER_DXA",
    "SPACING_AFTER_PT",
    "ROMAN_NUMERALS",
    "add_keep_next",
    "set_paragraph_spacing",
    "set_cell_valign_center",
    "clean_inherited_styles",
]
