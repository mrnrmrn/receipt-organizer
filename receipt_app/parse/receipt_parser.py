from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal

from receipt_app.config import DEFAULT_CONFIG
from receipt_app.models import OCRResult, ParsedReceipt, StructuredReceiptData


@dataclass
class ReceiptParser:
    allowed_pairs: tuple[tuple[str, str], ...]

    def parse(self, ocr_result: OCRResult) -> ParsedReceipt:
        structured = ocr_result.structured
        if structured is None:
            raise ValueError(
                "Gemini structured receipt data is missing. Plain-text fallback is disabled."
            )

        raw_text = self._resolve_raw_text(ocr_result.text, structured)
        category, subcategory = self._normalize_category_pair(
            structured.category,
            structured.subcategory,
        )

        return ParsedReceipt(
            source_file_name=ocr_result.source_file_name,
            raw_text=raw_text,
            amount=self._parse_amount(structured.amount),
            receipt_date=self._parse_date(structured.receipt_date),
            vendor=self._parse_vendor(structured.vendor),
            category=category,
            subcategory=subcategory,
            notes=self._parse_notes(structured.notes),
        )

    def _resolve_raw_text(self, text: str, structured: StructuredReceiptData) -> str:
        return (structured.raw_text or text).strip()

    def _parse_amount(self, amount: int | None) -> Decimal | None:
        if amount is None:
            return None
        return Decimal(str(amount))

    def _parse_date(self, receipt_date: str | None) -> date | None:
        if not receipt_date:
            return None
        try:
            return date.fromisoformat(receipt_date[:10])
        except ValueError:
            return None

    def _parse_vendor(self, vendor: str | None) -> str | None:
        if not vendor:
            return None
        return vendor.strip()[:60] or None

    def _normalize_category_pair(
        self,
        category: str | None,
        subcategory: str | None,
    ) -> tuple[str, str]:
        normalized_category = (category or "").strip()
        normalized_subcategory = (subcategory or "").strip()
        if (normalized_category, normalized_subcategory) in set(self.allowed_pairs):
            return (normalized_category, normalized_subcategory)
        return ("기타", "기타")

    def _parse_notes(self, notes: str | None) -> str | None:
        if not notes:
            return None
        return notes.strip() or None


def parse_receipt_text(ocr_result: OCRResult) -> ParsedReceipt:
    parser = ReceiptParser(
        allowed_pairs=DEFAULT_CONFIG.allowed_category_subcategory_pairs
    )
    return parser.parse(ocr_result)
