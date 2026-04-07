from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    template_path: Path = Path("최광현_2026-02-업무지원금신청.xlsx")
    sheet_name: str = "항목"
    operator_name_cell: str = "C1"
    total_amount_cell: str = "C2"
    month_cell: str = "C3"
    table_start_row: int = 5
    table_end_row: int = 24
    number_column: str = "B"
    category_column: str = "C"
    subcategory_column: str = "D"
    amount_column: str = "E"
    note_column: str = "F"
    image_anchor_columns: tuple[str, ...] = ("G", "H", "I", "J", "K")
    image_anchor_rows: tuple[int, ...] = (4, 31, 58, 85)
    image_end_row: int = 110
    max_receipt_images: int = 20
    tesseract_languages: str = "kor+eng"
    category_rules: dict[str, tuple[str, str]] = field(
        default_factory=lambda: {
            "택시": ("교통비", "택시"),
            "카카오택시": ("교통비", "택시"),
            "주차": ("교통비", "택시"),
            "버스": ("교통비", "대중교통"),
            "지하철": ("교통비", "대중교통"),
            "철도": ("교통비", "대중교통"),
            "ktx": ("교통비", "대중교통"),
            "티머니": ("교통비", "대중교통"),
            "tmoney": ("교통비", "대중교통"),
            "식대": ("복리후생비", "식비"),
            "식사": ("복리후생비", "식비"),
            "카페": ("복리후생비", "다과비"),
            "커피": ("복리후생비", "다과비"),
            "베이커리": ("복리후생비", "다과비"),
            "문구": ("사무용품비", "문구류"),
            "오피스": ("사무용품비", "사무용품"),
            "복사": ("사무용품비", "인쇄비"),
            "인쇄": ("사무용품비", "인쇄비"),
            "배송": ("운반비", "배송비"),
            "택배": ("운반비", "배송비"),
            "퀵": ("운반비", "배송비"),
            "소모품": ("소모품비", "일반소모품"),
            "편의점": ("소모품비", "일반소모품"),
            "통신": ("통신비", "통신비"),
            "u+": ("통신비", "통신비"),
            "lg u+": ("통신비", "통신비"),
            "kt": ("통신비", "통신비"),
            "skt": ("통신비", "통신비"),
        }
    )


DEFAULT_CONFIG = AppConfig()
