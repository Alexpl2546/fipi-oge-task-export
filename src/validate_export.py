from __future__ import annotations

import argparse
import csv
import json
import sqlite3
import sys
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import parse_qs, urlparse


@dataclass
class ValidationReport:
    records: int = 0
    malformed_lines: list[int] = field(default_factory=list)
    duplicate_qids: list[str] = field(default_factory=list)
    duplicate_urls: list[str] = field(default_factory=list)
    empty_text_qids: list[str] = field(default_factory=list)
    bad_url_qids: list[str] = field(default_factory=list)
    missing_image_paths: list[str] = field(default_factory=list)
    csv_count: int | None = None
    sqlite_count: int | None = None
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not (
            self.malformed_lines
            or self.duplicate_qids
            or self.duplicate_urls
            or self.empty_text_qids
            or self.bad_url_qids
            or self.missing_image_paths
            or self.errors
        )


def load_jsonl(path: Path, report: ValidationReport) -> list[dict]:
    records: list[dict] = []
    with path.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            if not line.strip():
                continue
            try:
                records.append(json.loads(line))
            except json.JSONDecodeError:
                report.malformed_lines.append(line_number)
    return records


def find_duplicates(values: list[str]) -> list[str]:
    seen: set[str] = set()
    duplicates: set[str] = set()
    for value in values:
        if value in seen:
            duplicates.add(value)
        seen.add(value)
    return sorted(duplicates)


def valid_task_url(url: str, qid: str | None) -> bool:
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    return (
        parsed.scheme in {"http", "https"}
        and parsed.netloc.endswith("fipi.ru")
        and parsed.path.endswith("/bank/index.php")
        and bool(query.get("proj", [""])[0])
        and (not qid or query.get("qid", [""])[0] == qid)
    )


def resolve_existing_path(raw_path: str, jsonl_path: Path) -> Path | None:
    candidate = Path(raw_path)
    candidates = [candidate]
    if not candidate.is_absolute():
        candidates.extend(
            [
                Path.cwd() / candidate,
                jsonl_path.parent / candidate,
                jsonl_path.parent.parent / candidate,
            ]
        )
    for item in candidates:
        if item.exists():
            return item
    return None


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


def validate(data_path: Path, csv_path: Path | None, sqlite_path: Path | None) -> ValidationReport:
    report = ValidationReport()
    if not data_path.exists():
        report.errors.append(f"JSONL file does not exist: {data_path}")
        return report

    records = load_jsonl(data_path, report)
    report.records = len(records)

    qids = [str(record.get("qid", "")) for record in records if record.get("qid")]
    urls = [str(record.get("url", "")) for record in records if record.get("url")]
    report.duplicate_qids = find_duplicates(qids)
    report.duplicate_urls = find_duplicates(urls)

    for record in records:
        qid = str(record.get("qid", ""))
        if not str(record.get("text", "")).strip():
            report.empty_text_qids.append(qid or "<missing qid>")
        url = str(record.get("url", ""))
        if not valid_task_url(url, qid):
            report.bad_url_qids.append(qid or url or "<missing url>")
        for image_path in record.get("local_images", []) or []:
            if resolve_existing_path(str(image_path), data_path) is None:
                report.missing_image_paths.append(str(image_path))

    csv_path = csv_path or data_path.with_suffix(".csv")
    sqlite_path = sqlite_path or data_path.with_suffix(".sqlite")
    report.csv_count = count_csv(csv_path)
    report.sqlite_count = count_sqlite(sqlite_path)
    if report.csv_count is not None and report.csv_count != report.records:
        report.errors.append(f"CSV row count mismatch: {report.csv_count} != {report.records}")
    if report.sqlite_count is not None and report.sqlite_count != report.records:
        report.errors.append(f"SQLite row count mismatch: {report.sqlite_count} != {report.records}")
    return report


def print_report(report: ValidationReport) -> None:
    print("Validation report")
    print(f"  JSONL records: {report.records}")
    print(f"  CSV records: {report.csv_count if report.csv_count is not None else 'not checked'}")
    print(f"  SQLite records: {report.sqlite_count if report.sqlite_count is not None else 'not checked'}")
    print(f"  Malformed JSONL lines: {len(report.malformed_lines)}")
    print(f"  Duplicate qids: {len(report.duplicate_qids)}")
    print(f"  Duplicate URLs: {len(report.duplicate_urls)}")
    print(f"  Empty texts: {len(report.empty_text_qids)}")
    print(f"  Bad URLs: {len(report.bad_url_qids)}")
    print(f"  Missing local images: {len(report.missing_image_paths)}")
    print(f"  Format/count errors: {len(report.errors)}")

    details = [
        ("malformed lines", report.malformed_lines[:10]),
        ("duplicate qids", report.duplicate_qids[:10]),
        ("duplicate URLs", report.duplicate_urls[:10]),
        ("empty text qids", report.empty_text_qids[:10]),
        ("bad URL qids", report.bad_url_qids[:10]),
        ("missing images", report.missing_image_paths[:10]),
        ("errors", report.errors[:10]),
    ]
    for label, values in details:
        if values:
            print(f"  First {label}: {values}")
    print("  Status:", "OK" if report.ok else "FAILED")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Validate a FIPI JSONL export.")
    parser.add_argument("--data", required=True, help="Path to tasks.jsonl.")
    parser.add_argument("--csv", default=None, help="Optional path to tasks.csv.")
    parser.add_argument("--sqlite", default=None, help="Optional path to tasks.sqlite.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    report = validate(
        Path(args.data),
        Path(args.csv) if args.csv else None,
        Path(args.sqlite) if args.sqlite else None,
    )
    print_report(report)
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
