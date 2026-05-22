from __future__ import annotations

import argparse
import csv
import json
import math
import sqlite3
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname
from PIL import Image as PILImage

try:
    from tqdm import tqdm
except ImportError:  # pragma: no cover
    tqdm = None


SOURCE_URL = "https://oge.fipi.ru/bank/index.php?proj=DE0E276E497AB3784C3FC4CC20248DC0"
TITLE = "ФИПИ ОГЭ Математика — полный экспорт заданий"
MAX_IMAGE_WIDTH_PX = 680
MAX_CELL_TEXT = 32000


@dataclass
class ExportStats:
    tasks_processed: int = 0
    cards_created: int = 0
    images_embedded: int = 0
    images_linked: int = 0
    images_skipped: int = 0
    accessible_images: int = 0
    errors: list[dict[str, str]] = field(default_factory=list)


def load_tasks(path: Path, limit: int | None) -> list[dict[str, Any]]:
    tasks: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            if not line.strip():
                continue
            task = json.loads(line)
            task["_line_number"] = line_number
            tasks.append(task)
            if limit and len(tasks) >= limit:
                break
    return tasks


def read_all_jsonl_count(path: Path) -> int:
    count = 0
    with path.open("r", encoding="utf-8") as file:
        for line in file:
            if line.strip():
                count += 1
    return count


def count_csv(path: Path) -> int | None:
    if not path.exists():
        return None
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return sum(1 for _ in csv.DictReader(file))


def count_sqlite(path: Path) -> int | None:
    if not path.exists():
        return None
    conn = sqlite3.connect(path)
    try:
        return int(conn.execute("SELECT COUNT(*) FROM tasks").fetchone()[0])
    finally:
        conn.close()


def parse_validation_status(report_path: Path) -> str:
    if not report_path.exists():
        return "not found"
    text = report_path.read_text(encoding="utf-8", errors="replace")
    return "OK" if "Status: OK" in text or "validation_status | OK" in text else "unknown"


def resolve_path(raw_path: str, data_path: Path) -> Path | None:
    candidate = Path(raw_path)
    candidates = [candidate]
    if not candidate.is_absolute():
        candidates.extend(
            [
                Path.cwd() / candidate,
                data_path.parent / candidate,
                data_path.parent.parent / candidate,
            ]
        )
    for item in candidates:
        if item.exists():
            return item.resolve()
    return None


def markdown_path_for(task: dict[str, Any], data_path: Path) -> Path:
    return data_path.parent / "markdown" / f"{task.get('qid')}.md"


def make_short_text(text: str, limit: int = 280) -> str:
    text = " ".join(str(text or "").split())
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def safe_text(text: Any, stats: ExportStats, qid: str, field: str) -> str:
    value = "" if text is None else str(text)
    if len(value) > MAX_CELL_TEXT:
        stats.errors.append(
            {"qid": qid, "type": "too_long_text", "message": f"{field} truncated from {len(value)} chars"}
        )
        return value[:MAX_CELL_TEXT]
    return value


def options_text(options: list[dict[str, Any]]) -> str:
    lines = []
    for option in options or []:
        value = str(option.get("value", "")).strip()
        text = str(option.get("text", "")).strip()
        if value and text and value != text:
            lines.append(f"{value}: {text}")
        else:
            lines.append(text or value)
    return "\n".join(lines)


def set_common_layout(wb: Workbook) -> None:
    wb.properties.creator = "Codex"
    wb.properties.title = TITLE
    wb.properties.subject = "FIPI OGE math task export"


def style_header(row) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in row:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def create_summary_sheet(
    wb: Workbook,
    data_path: Path,
    tasks: list[dict[str, Any]],
    total_jsonl_count: int,
    mode: str,
) -> None:
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 90

    root = data_path.parent
    values = [
        ("Название выгрузки", TITLE),
        ("Исходный URL", SOURCE_URL),
        ("Дата формирования Excel", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Режим Excel", mode),
        ("Количество заданий в Excel", len(tasks)),
        ("Количество изображений в Excel JSONL slice", sum(len(t.get("local_images") or []) for t in tasks)),
        ("Количество записей JSONL", total_jsonl_count),
        ("Количество записей CSV", count_csv(root / "tasks.csv")),
        ("Количество записей SQLite", count_sqlite(root / "tasks.sqlite")),
        ("Статус последней валидации", parse_validation_status(root / "export_report.md")),
        (
            "Правильные ответы",
            "Открытый HTML ФИПИ не выдаёт правильные ответы для выгруженных заданий; пустое поле answer сохраняется как null.",
        ),
        ("Примечание о правах", "Права на задания и материалы сохраняются за ФИПИ/правообладателями."),
    ]
    ws["A1"] = "Параметр"
    ws["B1"] = "Значение"
    style_header(ws[1])
    for row_idx, (name, value) in enumerate(values, start=2):
        ws.cell(row_idx, 1, name)
        ws.cell(row_idx, 2, "" if value is None else value)
        ws.cell(row_idx, 1).font = Font(bold=True, color="374151")
        ws.cell(row_idx, 1).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row_idx, 2).alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"


def create_index_sheet(wb: Workbook, tasks: list[dict[str, Any]], data_path: Path) -> None:
    ws = wb.create_sheet("Index")
    headers = [
        "№",
        "qid",
        "URL",
        "КЭС / раздел",
        "Тема",
        "Номер",
        "Краткий текст",
        "Изображений",
        "Карточка",
        "Markdown",
    ]
    ws.append(headers)
    style_header(ws[1])
    widths = [8, 14, 36, 42, 22, 14, 70, 12, 18, 36]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{len(tasks) + 1}"

    for row_idx, task in enumerate(tasks, start=2):
        task_no = row_idx - 1
        qid = str(task.get("qid", ""))
        markdown_path = markdown_path_for(task, data_path)
        kes = "; ".join(task.get("kes") or [])
        values = [
            task_no,
            qid,
            task.get("url", ""),
            kes,
            task.get("theme") or "",
            task.get("task_number") or qid,
            make_short_text(task.get("text", "")),
            len(task.get("local_images") or []),
            "Открыть карточку",
            str(markdown_path),
        ]
        ws.append(values)
        excel_row = row_idx
        task_anchor = task.get("_excel_task_row")
        if task_anchor:
            ws.cell(excel_row, 9).hyperlink = f"#{quote_sheetname('Tasks')}!A{task_anchor}"
            ws.cell(excel_row, 9).style = "Hyperlink"
        ws.cell(excel_row, 2).hyperlink = task.get("url", "")
        ws.cell(excel_row, 2).style = "Hyperlink"
        ws.cell(excel_row, 3).hyperlink = task.get("url", "")
        ws.cell(excel_row, 3).style = "Hyperlink"
        if markdown_path.exists():
            ws.cell(excel_row, 10).hyperlink = markdown_path.resolve().as_uri()
            ws.cell(excel_row, 10).style = "Hyperlink"
        for cell in ws[excel_row]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_row=len(tasks) + 1):
        ws.row_dimensions[row[0].row].height = 45


def setup_tasks_sheet(wb: Workbook) -> Any:
    ws = wb.create_sheet("Tasks")
    widths = [6, 14, 18, 18, 18, 18, 18, 18, 18, 18]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A1"
    return ws


def merge_write(ws, row: int, start_col: int, end_col: int, value: str, style: str = "body") -> None:
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = ws.cell(row, start_col, value)
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    if style == "title":
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.font = Font(bold=True, size=12, color="1F4E78")
    elif style == "service":
        cell.fill = PatternFill("solid", fgColor="F3F4F6")
        cell.font = Font(size=9, color="6B7280")
    else:
        cell.font = Font(size=11, color="111827")


def apply_card_border(ws, start_row: int, end_row: int, start_col: int = 1, end_col: int = 10) -> None:
    side = Side(style="thin", color="D1D5DB")
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row, col)
            cell.border = Border(
                left=side if col == start_col else cell.border.left,
                right=side if col == end_col else cell.border.right,
                top=side if row == start_row else cell.border.top,
                bottom=side if row == end_row else cell.border.bottom,
            )


def image_size(path: Path) -> tuple[int, int]:
    with PILImage.open(path) as image:
        return image.size


def scaled_size(width: int, height: int, max_width: int = MAX_IMAGE_WIDTH_PX) -> tuple[int, int]:
    if width <= max_width:
        return width, height
    scale = max_width / width
    return int(width * scale), int(height * scale)


def build_tasks_sheet(
    ws,
    tasks: list[dict[str, Any]],
    data_path: Path,
    mode: str,
    stats: ExportStats,
) -> None:
    iterable = tasks
    if tqdm:
        iterable = tqdm(tasks, desc=f"Excel {mode}", unit="task")
    row = 1
    for idx, task in enumerate(iterable, start=1):
        qid = str(task.get("qid", ""))
        card_start = row
        task["_excel_task_row"] = row
        try:
            text = safe_text(task.get("text", ""), stats, qid, "text")
            if not text.strip():
                stats.errors.append({"qid": qid, "type": "empty_text", "message": "Task text is empty"})
            title = f"Задание № {idx} / {qid} / источник ФИПИ"
            merge_write(ws, row, 1, 10, title, "title")
            ws.cell(row, 1).hyperlink = task.get("url", "")
            ws.cell(row, 1).font = Font(bold=True, size=12, color="0563C1", underline="single")
            ws.cell(row, 1).fill = PatternFill("solid", fgColor="D9EAF7")
            ws.row_dimensions[row].height = 26
            row += 1

            merge_write(ws, row, 1, 10, text)
            ws.row_dimensions[row].height = min(260, max(45, math.ceil(len(text) / 95) * 16))
            row += 1

            opt_text = options_text(task.get("options") or [])
            if opt_text:
                opt_text = safe_text("Варианты ответа:\n" + opt_text, stats, qid, "options")
                merge_write(ws, row, 1, 10, opt_text)
                ws.row_dimensions[row].height = min(220, max(35, math.ceil(len(opt_text) / 90) * 16))
                row += 1

            local_images = task.get("local_images") or []
            if local_images:
                merge_write(ws, row, 1, 10, "Изображения / схемы / графики", "service")
                ws.row_dimensions[row].height = 20
                row += 1
            for img_idx, raw_path in enumerate(local_images, start=1):
                image_path = resolve_path(str(raw_path), data_path)
                if image_path is None:
                    stats.images_skipped += 1
                    stats.errors.append({"qid": qid, "type": "missing_image", "message": str(raw_path)})
                    merge_write(ws, row, 1, 10, f"Изображение не найдено: {raw_path}", "service")
                    row += 1
                    continue
                stats.accessible_images += 1
                if mode == "links":
                    merge_write(ws, row, 1, 10, f"Изображение {img_idx}: {image_path}", "service")
                    ws.cell(row, 1).hyperlink = image_path.as_uri()
                    ws.cell(row, 1).style = "Hyperlink"
                    ws.row_dimensions[row].height = 22
                    stats.images_linked += 1
                    row += 1
                    continue
                try:
                    width, height = image_size(image_path)
                    scaled_w, scaled_h = scaled_size(width, height)
                    img = XLImage(str(image_path))
                    img.width = scaled_w
                    img.height = scaled_h
                    ws.add_image(img, f"B{row}")
                    ws.row_dimensions[row].height = max(24, scaled_h * 0.75 + 8)
                    stats.images_embedded += 1
                    row += 1
                except Exception as exc:
                    stats.images_skipped += 1
                    stats.errors.append({"qid": qid, "type": "image_read_error", "message": f"{image_path}: {exc}"})
                    merge_write(ws, row, 1, 10, f"Не удалось встроить изображение: {image_path}", "service")
                    row += 1

            markdown_path = markdown_path_for(task, data_path)
            service = (
                f"qid: {qid}\n"
                f"URL: {task.get('url', '')}\n"
                f"Локальные изображения: {json.dumps(local_images, ensure_ascii=False)}\n"
                f"Markdown: {markdown_path}"
            )
            service = safe_text(service, stats, qid, "service")
            merge_write(ws, row, 1, 10, service, "service")
            ws.row_dimensions[row].height = min(120, max(45, math.ceil(len(service) / 110) * 14))
            row += 1

            apply_card_border(ws, card_start, row - 1)
            row += 2
            stats.cards_created += 1
            stats.tasks_processed += 1
        except Exception as exc:
            stats.errors.append({"qid": qid, "type": "task_error", "message": str(exc)})
            row = max(row + 2, card_start + 3)


def create_errors_sheet(wb: Workbook, errors: list[dict[str, str]]) -> None:
    if not errors:
        return
    ws = wb.create_sheet("Errors")
    ws.append(["qid", "type", "message"])
    style_header(ws[1])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 100
    for error in errors:
        ws.append([error.get("qid"), error.get("type"), error.get("message")])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:C{len(errors) + 1}"
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def order_sheets(wb: Workbook) -> None:
    desired = ["Summary", "Index", "Tasks", "Errors"]
    wb._sheets.sort(key=lambda ws: desired.index(ws.title) if ws.title in desired else len(desired))


def validate_workbook(path: Path, tasks: list[dict[str, Any]], stats: ExportStats, mode: str) -> list[str]:
    issues: list[str] = []
    wb = load_workbook(path, read_only=False)
    try:
        index = wb["Index"]
        tasks_sheet = wb["Tasks"]
        index_qids = [index.cell(row=i, column=2).value for i in range(2, index.max_row + 1)]
        if len(index_qids) != len(tasks):
            issues.append(f"Index qid count mismatch: {len(index_qids)} != {len(tasks)}")
        if len(set(index_qids)) != len(index_qids):
            issues.append("Duplicate qids found in Index sheet")
        card_count = 0
        for row in range(1, tasks_sheet.max_row + 1):
            cell = tasks_sheet.cell(row=row, column=1)
            value = str(cell.value or "")
            fill = cell.fill.fgColor.rgb
            if value.startswith("Задание") and fill in {"00D9EAF7", "FFD9EAF7"}:
                card_count += 1
        if card_count != len(tasks):
            issues.append(f"Task card count mismatch: {card_count} != {len(tasks)}")
        for task in tasks:
            for raw_path in task.get("local_images") or []:
                if resolve_path(str(raw_path), Path(path).parent / "tasks.jsonl") is None:
                    issues.append(f"Missing local image link: {raw_path}")
        if mode == "embedded" and stats.images_embedded != stats.accessible_images:
            issues.append(f"Embedded image mismatch: {stats.images_embedded} != {stats.accessible_images}")
    finally:
        wb.close()
    return issues


def build_excel(args: argparse.Namespace) -> ExportStats:
    data_path = Path(args.data)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tasks = load_tasks(data_path, args.limit)
    total_jsonl_count = read_all_jsonl_count(data_path)
    stats = ExportStats()

    wb = Workbook()
    set_common_layout(wb)
    create_summary_sheet(wb, data_path, tasks, total_jsonl_count, args.mode)
    tasks_sheet = setup_tasks_sheet(wb)
    build_tasks_sheet(tasks_sheet, tasks, data_path, args.mode, stats)
    create_index_sheet(wb, tasks, data_path)
    create_errors_sheet(wb, stats.errors)
    order_sheets(wb)
    wb.save(out_path)

    issues = validate_workbook(out_path, tasks, stats, args.mode)
    for issue in issues:
        stats.errors.append({"qid": "", "type": "workbook_validation", "message": issue})
    if issues:
        # Persist validation issues to the workbook after initial save.
        wb = load_workbook(out_path)
        if "Errors" in wb.sheetnames:
            ws = wb["Errors"]
        else:
            ws = wb.create_sheet("Errors")
            ws.append(["qid", "type", "message"])
            style_header(ws[1])
        for issue in issues:
            ws.append(["", "workbook_validation", issue])
        order_sheets(wb)
        wb.save(out_path)
        wb.close()
    return stats


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create formatted Excel workbooks from FIPI JSONL export.")
    parser.add_argument("--data", required=True, help="Path to tasks.jsonl.")
    parser.add_argument("--out", required=True, help="Output .xlsx path.")
    parser.add_argument("--mode", choices=("embedded", "links"), default="links", help="Embed images or link to them.")
    parser.add_argument("--limit", type=int, default=None, help="Limit number of tasks for preview.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    stats = build_excel(args)
    out_path = Path(args.out)
    size_mb = out_path.stat().st_size / (1024 * 1024) if out_path.exists() else 0
    print("Excel export report")
    print(f"  tasks processed: {stats.tasks_processed}")
    print(f"  cards created: {stats.cards_created}")
    print(f"  images embedded: {stats.images_embedded}")
    print(f"  images linked: {stats.images_linked}")
    print(f"  images skipped: {stats.images_skipped}")
    print(f"  errors: {len(stats.errors)}")
    print(f"  xlsx size: {size_mb:.2f} MB")
    if args.mode == "embedded" and size_mb > 200:
        print("  warning: embedded workbook is large and may open slowly")
    return 0 if not stats.errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
