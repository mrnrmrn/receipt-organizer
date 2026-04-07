from __future__ import annotations

from datetime import date
from decimal import Decimal
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

from receipt_app.config import AppConfig, DEFAULT_CONFIG
from receipt_app.models import ExportRow, UploadedReceipt
from receipt_app.utils.images import (
    image_to_png_bytes,
    open_image_from_bytes,
    resize_for_excel,
)


def build_workbook_bytes(
    person_name: str,
    rows: list[ExportRow | dict],
    receipts: list[UploadedReceipt],
    template_path: str | Path | None = None,
    month: date | str | None = None,
    config: AppConfig = DEFAULT_CONFIG,
) -> bytes:
    normalized_rows = _coerce_rows(rows)
    report_month = month or date.today().replace(day=1)
    return export_receipts_to_workbook(
        rows=normalized_rows,
        receipts=receipts,
        operator_name=person_name,
        report_month_text=report_month,
        template_path=template_path,
        config=config,
    )


def export_receipts_to_workbook(
    rows: list[ExportRow],
    receipts: list[UploadedReceipt],
    operator_name: str,
    report_month_text: str,
    template_path: str | Path | None = None,
    config: AppConfig = DEFAULT_CONFIG,
) -> bytes:
    workbook = load_workbook(filename=Path(template_path or config.template_path))
    sheet = workbook[config.sheet_name]

    sheet[config.operator_name_cell] = operator_name
    sheet[config.month_cell] = report_month_text
    total_amount = sum((row.amount for row in rows), start=0)
    sheet[config.total_amount_cell] = int(total_amount)

    _write_rows(sheet=sheet, rows=rows, config=config)
    _embed_receipt_images(sheet=sheet, receipts=receipts, config=config)

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _write_rows(sheet, rows: list[ExportRow], config: AppConfig) -> None:
    for row_number in range(config.table_start_row, config.table_end_row + 1):
        sheet[f"{config.number_column}{row_number}"] = None
        sheet[f"{config.category_column}{row_number}"] = None
        sheet[f"{config.subcategory_column}{row_number}"] = None
        sheet[f"{config.amount_column}{row_number}"] = None
        sheet[f"{config.note_column}{row_number}"] = None

    for index, export_row in enumerate(rows, start=1):
        row_number = config.table_start_row + index - 1
        if row_number > config.table_end_row:
            break
        sheet[f"{config.number_column}{row_number}"] = index
        sheet[f"{config.category_column}{row_number}"] = export_row.category
        sheet[f"{config.subcategory_column}{row_number}"] = export_row.subcategory
        sheet[f"{config.amount_column}{row_number}"] = int(export_row.amount)
        sheet[f"{config.note_column}{row_number}"] = export_row.note_text


def _embed_receipt_images(
    sheet, receipts: list[UploadedReceipt], config: AppConfig
) -> None:
    if hasattr(sheet, "_images"):
        sheet._images = []

    if not receipts:
        return

    slots = [
        f"{column}{row}"
        for row in config.image_anchor_rows
        for column in config.image_anchor_columns
    ]
    for receipt, anchor in zip(receipts[: config.max_receipt_images], slots):
        image = open_image_from_bytes(receipt.image_bytes)
        resized = resize_for_excel(image, max_width=150, max_height=170)
        excel_image = XLImage(BytesIO(image_to_png_bytes(resized)))
        excel_image.anchor = anchor
        sheet.add_image(excel_image)


def _coerce_rows(rows: list[ExportRow | dict]) -> list[ExportRow]:
    normalized: list[ExportRow] = []
    for index, row in enumerate(rows, start=1):
        if isinstance(row, ExportRow):
            normalized.append(row)
            continue

        amount_value = row.get("amount")
        if amount_value in (None, ""):
            continue

        normalized.append(
            ExportRow(
                number=index,
                category=(row.get("category") or "기타").strip(),
                subcategory=(row.get("subcategory") or "기타").strip(),
                amount=Decimal(str(amount_value).replace(",", "").strip()),
                vendor=(row.get("vendor") or None),
                receipt_date=row.get("receipt_date"),
                notes=(row.get("notes") or None),
            )
        )

    return normalized
