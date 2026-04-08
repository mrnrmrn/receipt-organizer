from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass
from typing import Optional, cast

from pydantic import BaseModel, ValidationError
import streamlit as st

from receipt_app.config import DEFAULT_CONFIG
from receipt_app.models import OCRResult, StructuredReceiptData, UploadedReceipt
from receipt_app.utils.images import (
    image_to_png_bytes,
    normalize_receipt_image,
    open_image_from_bytes,
)


def _build_default_prompt() -> str:
    allowed_pairs = "\n".join(
        f"- {category} / {subcategory}"
        for category, subcategory in DEFAULT_CONFIG.allowed_category_subcategory_pairs
    )
    return (
        "당신은 영수증 구조화 추출기입니다.\n"
        "입력 이미지는 한국어 영수증 또는 결제내역 캡처입니다.\n"
        "목표:\n"
        "- 아래 스키마에 맞는 JSON만 반환합니다.\n"
        "- 값이 확실하지 않으면 null을 사용합니다.\n"
        "- 추측하거나 만들어내지 않습니다.\n"
        "- 설명, 마크다운, 코드블록 없이 JSON만 반환합니다.\n"
        "추출 규칙:\n"
        "1. amount는 실제 결제/청구 총액만 반환합니다.\n"
        "2. receipt_date는 영수증에 보이는 실제 결제일을 우선 사용하고 YYYY-MM-DD 형식으로 반환합니다.\n"
        "3. vendor는 가맹점/상호명을 짧고 자연스럽게 반환합니다.\n"
        "4. raw_text에는 영수증에서 읽힌 주요 텍스트를 줄바꿈 포함해서 넣습니다.\n"
        "5. category, subcategory는 아래 허용 목록 중 하나만 선택합니다.\n"
        "6. 허용 목록으로 분류가 어렵다면 category와 subcategory는 둘 다 '기타'를 반환합니다.\n"
        f"{DEFAULT_CONFIG.structured_extraction_guide_text}\n"
        "허용 category/subcategory 조합:\n"
        f"{allowed_pairs}"
    )


DEFAULT_GEMINI_PROMPT = _build_default_prompt()


class GeminiStructuredReceipt(BaseModel):
    raw_text: Optional[str] = None
    amount: Optional[int] = None
    receipt_date: Optional[str] = None
    vendor: Optional[str] = None
    category: Optional[str] = None
    subcategory: Optional[str] = None
    notes: Optional[str] = None


def _read_streamlit_secret(name: str) -> str | None:
    try:
        value = st.secrets.get(name)
    except Exception:
        return None

    if isinstance(value, str) and value.strip():
        return value.strip()
    return None


def _read_streamlit_section_secret(section: str, key: str) -> str | None:
    try:
        value = st.secrets.get(section)
    except Exception:
        return None

    if isinstance(value, Mapping):
        nested_value = cast(Mapping[str, object], value).get(key)
        if isinstance(nested_value, str) and nested_value.strip():
            return nested_value.strip()
    return None


def _get_server_setting(
    *names: str, section: str | None = None, key: str | None = None
) -> str | None:
    for name in names:
        value = _read_streamlit_secret(name)
        if value:
            return value

    if section and key:
        value = _read_streamlit_section_secret(section, key)
        if value:
            return value

    for name in names:
        value = os.getenv(name)
        if value and value.strip():
            return value.strip()

    return None


def _collect_response_text(response: object) -> str:
    text = getattr(response, "text", None)
    if isinstance(text, str) and text.strip():
        return text.strip()

    parts_text: list[str] = []
    candidates = cast(list[object], getattr(response, "candidates", None) or [])
    for candidate in candidates:
        content = getattr(candidate, "content", None)
        parts = cast(list[object], getattr(content, "parts", None) or [])
        for part in parts:
            part_text = getattr(part, "text", None)
            if isinstance(part_text, str) and part_text.strip():
                parts_text.append(part_text.strip())

    return "\n".join(parts_text).strip()


def _parse_structured_response(response: object) -> GeminiStructuredReceipt:
    parsed = getattr(response, "parsed", None)
    if isinstance(parsed, GeminiStructuredReceipt):
        return parsed
    if parsed is not None:
        return GeminiStructuredReceipt.model_validate(parsed)

    text = _collect_response_text(response)
    if not text:
        raise ValueError("Gemini returned an empty structured response.")

    try:
        return GeminiStructuredReceipt.model_validate_json(text)
    except ValidationError as exc:
        raise ValueError("Gemini returned invalid structured receipt JSON.") from exc


def _to_structured_receipt_data(
    structured: GeminiStructuredReceipt,
) -> StructuredReceiptData:
    raw_text = (structured.raw_text or "").strip()
    return StructuredReceiptData(
        raw_text=raw_text,
        amount=structured.amount,
        receipt_date=(structured.receipt_date or "").strip() or None,
        vendor=(structured.vendor or "").strip() or None,
        category=(structured.category or "").strip() or None,
        subcategory=(structured.subcategory or "").strip() or None,
        notes=(structured.notes or "").strip() or None,
    )


@dataclass
class GeminiOCRBackend:
    model: str | None = None
    prompt: str = DEFAULT_GEMINI_PROMPT
    backend_name: str = "gemini"
    language: str = "ko,en"

    def extract_text(self, receipt: UploadedReceipt) -> OCRResult:
        from google import genai
        from google.genai import types

        api_key = _get_server_setting(
            "GEMINI_API_KEY",
            "GOOGLE_API_KEY",
            section="gemini",
            key="api_key",
        )
        if not api_key:
            raise ValueError(
                "Gemini API key is not configured. Set GEMINI_API_KEY in Streamlit secrets or the server environment."
            )

        model = self.model or _get_server_setting(
            "GEMINI_MODEL",
            section="gemini",
            key="model",
        )
        model = model or getattr(DEFAULT_CONFIG, "gemini_model", "gemini-2.5-flash")

        image = open_image_from_bytes(receipt.image_bytes)
        normalized = normalize_receipt_image(image)
        png_bytes = image_to_png_bytes(normalized)

        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=model,
            contents=[
                types.Content(
                    role="user",
                    parts=[
                        types.Part.from_text(text=self.prompt),
                        types.Part.from_bytes(data=png_bytes, mime_type="image/png"),
                    ],
                )
            ],
            config=types.GenerateContentConfig(
                temperature=0,
                response_mime_type="application/json",
                response_schema=GeminiStructuredReceipt,
            ),
        )

        structured = _to_structured_receipt_data(_parse_structured_response(response))
        text = structured.raw_text or _collect_response_text(response)
        if not text:
            raise ValueError("Gemini returned an empty OCR response.")

        lines = [line.strip() for line in text.splitlines() if line.strip()]
        return OCRResult(
            source_file_name=receipt.file_name,
            text=text,
            backend_name=self.backend_name,
            language=self.language,
            lines=lines,
            structured=structured,
        )
