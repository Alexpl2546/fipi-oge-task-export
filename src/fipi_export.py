from __future__ import annotations

import argparse
import copy
import csv
import hashlib
import json
import logging
import math
import os
import re
import sqlite3
import ssl
import sys
import time
import warnings
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlencode, urljoin, urlparse
from urllib.request import Request, urlopen

from lxml import html

try:
    import certifi
except ImportError:  # pragma: no cover
    certifi = None

try:
    import truststore
except ImportError:  # pragma: no cover
    truststore = None

try:
    import requests
except ImportError:  # pragma: no cover - fallback keeps the tool runnable in lean envs.
    requests = None

try:
    from urllib3.exceptions import InsecureRequestWarning
except ImportError:  # pragma: no cover
    InsecureRequestWarning = None

try:
    from tqdm import tqdm
except ImportError:  # pragma: no cover
    tqdm = None


DEFAULT_URL = "https://oge.fipi.ru/bank/index.php?proj=DE0E276E497AB3784C3FC4CC20248DC0"
USER_AGENT = "fipi-local-exporter/0.1 (+local educational use; single-threaded)"
PROTECTION_MARKERS = (
    "captcha",
    "access denied",
    "доступ ограничен",
    "проверка безопасности",
    "cloudflare",
)


class ExporterError(RuntimeError):
    pass


@dataclass
class FetchResult:
    url: str
    status_code: int
    content: bytes
    content_type: str


@dataclass
class ExportStats:
    found: int = 0
    saved: int = 0
    skipped: int = 0
    duplicates: int = 0
    empty_text: int = 0
    images_downloaded: int = 0
    image_errors: int = 0
    task_errors: int = 0
    pages_fetched: int = 0


class FipiClient:
    def __init__(self, delay: float, timeout: float, debug_dir: Path | None, tls_mode: str = "certifi") -> None:
        self.delay = delay
        self.timeout = timeout
        self.debug_dir = debug_dir
        self.last_request_at = 0.0
        self.tls_mode = tls_mode
        self.insecure = tls_mode == "insecure"
        self.backend = "urllib" if tls_mode == "system" else ("requests" if requests else "urllib")
        self.certifi_path = certifi.where() if certifi else None
        self.session = None
        self.ssl_context: ssl.SSLContext | None = None
        if self.tls_mode == "system":
            if truststore is None:
                raise ExporterError(
                    "TLS mode 'system' requires the truststore package. Install dependencies with "
                    "`pip install -r requirements.txt`."
                )
            truststore.inject_into_ssl()
        if self.backend == "requests":
            self.session = requests.Session()
            self.session.headers.update({"User-Agent": USER_AGENT})
            if self.insecure and InsecureRequestWarning:
                warnings.simplefilter("ignore", InsecureRequestWarning)
                requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)
        else:
            if self.insecure:
                self.ssl_context = ssl._create_unverified_context()
            elif self.tls_mode == "system":
                self.ssl_context = ssl.create_default_context()
            elif self.certifi_path:
                self.ssl_context = ssl.create_default_context(cafile=self.certifi_path)
            else:
                self.ssl_context = ssl.create_default_context()

    @property
    def verify_description(self) -> str:
        if self.insecure:
            return "verify=False / unverified SSLContext"
        if self.tls_mode == "system":
            return "Windows/system trust store via truststore"
        if self.backend == "requests":
            return f"verify={self.certifi_path}" if self.certifi_path else "verify=True (certifi unavailable)"
        return f"SSLContext(cafile={self.certifi_path})" if self.certifi_path else "SSLContext(default)"

    def log_configuration(self) -> None:
        logging.info("Python version: %s", sys.version.replace("\n", " "))
        logging.info("certifi.where(): %s", self.certifi_path or "certifi is not installed")
        logging.info("truststore: %s", "installed" if truststore else "not installed")
        logging.info("tls_mode: %s", self.tls_mode)
        logging.info("insecure: %s", str(self.insecure).lower())
        logging.info("HTTP backend: %s", self.backend)
        logging.info("TLS verify/context: %s", self.verify_description)

    def sleep_if_needed(self) -> None:
        elapsed = time.monotonic() - self.last_request_at
        if elapsed < self.delay:
            time.sleep(self.delay - elapsed)

    def fetch(self, url: str, *, referer: str | None = None, binary: bool = False) -> FetchResult:
        headers = {"User-Agent": USER_AGENT}
        if referer:
            headers["Referer"] = referer
        for attempt in range(2):
            self.sleep_if_needed()
            try:
                if self.backend == "requests":
                    verify: bool | str
                    if self.insecure:
                        verify = False
                    elif self.certifi_path:
                        verify = self.certifi_path
                    else:
                        verify = True
                    response = self.session.get(
                        url,
                        headers=headers,
                        timeout=self.timeout,
                        verify=verify,
                    )
                    content = response.content
                    result = FetchResult(url, response.status_code, content, response.headers.get("content-type", ""))
                else:
                    req = Request(url, headers=headers)
                    with urlopen(req, timeout=self.timeout, context=self.ssl_context) as response:
                        result = FetchResult(
                            response.geturl(),
                            int(response.status),
                            response.read(),
                            response.headers.get("content-type", ""),
                        )
            except HTTPError as exc:
                content = exc.read()
                result = FetchResult(url, exc.code, content, exc.headers.get("content-type", ""))
            except ssl.SSLError as exc:
                raise ExporterError(
                    "TLS certificate verification failed. Check the OS/Python trust store, install/update certifi, "
                    "or check whether a corporate TLS inspection certificate must be trusted. Rerun with --insecure "
                    "only as a diagnostic step."
                ) from exc
            except URLError as exc:
                if isinstance(exc.reason, ssl.SSLError):
                    raise ExporterError(
                        "TLS certificate verification failed. Check the OS/Python trust store, install/update certifi, "
                        "or check whether a corporate TLS inspection certificate must be trusted. Rerun with --insecure "
                        "only as a diagnostic step."
                    ) from exc
                raise ExporterError(f"Network error while requesting {url}: {exc.reason}") from exc
            except Exception as exc:
                if requests and isinstance(exc, requests.exceptions.SSLError):
                    raise ExporterError(
                        "TLS certificate verification failed. Check the OS/Python trust store, install/update certifi, "
                        "or check whether a corporate TLS inspection certificate must be trusted. Rerun with --insecure "
                        "only as a diagnostic step."
                    ) from exc
                raise ExporterError(f"Network error while requesting {url}: {exc}") from exc
            finally:
                self.last_request_at = time.monotonic()

            if result.status_code == 429 and attempt == 0:
                self.delay = max(self.delay * 2, 5)
                logging.warning("FIPI returned HTTP 429. Increasing delay to %.1fs and retrying once.", self.delay)
                time.sleep(self.delay)
                continue
            break

        if result.status_code in {403, 404, 429, 500, 502, 503, 504}:
            raise ExporterError(f"FIPI returned HTTP {result.status_code} for {url}. Stopping export.")
        if not binary:
            text_probe = decode_html(result.content).lower()
            if any(marker in text_probe for marker in PROTECTION_MARKERS):
                raise ExporterError(
                    "The response looks like a protection or access-limitation page. Stopping without bypass attempts."
                )
            self.save_debug_page(url, result.content)
        return result

    def save_debug_page(self, url: str, content: bytes) -> None:
        if not self.debug_dir:
            return
        self.debug_dir.mkdir(parents=True, exist_ok=True)
        name = hashlib.sha1(url.encode("utf-8")).hexdigest()[:12] + ".html"
        (self.debug_dir / name).write_bytes(content)


def diagnose_tls(url: str, timeout: float) -> int:
    backend = "requests" if requests else "urllib"
    cert_path = certifi.where() if certifi else None
    print("TLS diagnostics")
    print(f"  Python version: {sys.version.replace(chr(10), ' ')}")
    print(f"  HTTP backend: {backend}")
    print(f"  certifi.where(): {cert_path or 'certifi is not installed'}")
    print(f"  truststore: {'installed' if truststore else 'not installed'}")
    print(f"  SSL_CERT_FILE: {os.environ.get('SSL_CERT_FILE') or '<unset>'}")
    print(f"  REQUESTS_CA_BUNDLE: {os.environ.get('REQUESTS_CA_BUNDLE') or '<unset>'}")
    print(f"  CURL_CA_BUNDLE: {os.environ.get('CURL_CA_BUNDLE') or '<unset>'}")
    print(f"  URL: {url}")

    modes = ["default", "certifi", "system", "insecure"]
    results: list[tuple[str, bool, str]] = []
    for mode in modes:
        ok, message = diagnose_tls_mode(url, timeout, backend, mode, cert_path)
        results.append((mode, ok, message))
        status = "OK" if ok else "FAILED"
        print(f"  {mode}: {status} - {message}")

    if any(ok for mode, ok, _ in results if mode != "insecure"):
        print("  Conclusion: TLS verification works without --insecure.")
        return 0
    if any(ok for mode, ok, _ in results if mode == "insecure"):
        print(
            "  Conclusion: only --insecure works. Check local trust store, certifi installation/path, "
            "or corporate TLS inspection certificates."
        )
        return 1
    print("  Conclusion: all modes failed; this may be a network or site availability issue.")
    return 2


def diagnose_tls_mode(url: str, timeout: float, backend: str, mode: str, cert_path: str | None) -> tuple[bool, str]:
    try:
        if backend == "requests":
            if mode == "default":
                verify: bool | str = True
            elif mode == "certifi":
                if not cert_path:
                    return False, "certifi is not installed"
                verify = cert_path
            elif mode == "system":
                if truststore is None:
                    return False, "truststore is not installed"
                truststore.inject_into_ssl()
                verify = True
            else:
                verify = False
                if InsecureRequestWarning:
                    warnings.simplefilter("ignore", InsecureRequestWarning)
                    requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)
            session = requests.Session()
            session.headers.update({"User-Agent": USER_AGENT})
            response = session.get(url, timeout=timeout, verify=verify)
            response.close()
            return True, f"HTTP {response.status_code}; verify={verify}"

        if mode == "default":
            context = ssl.create_default_context()
            context_info = "default SSLContext"
        elif mode == "certifi":
            if not cert_path:
                return False, "certifi is not installed"
            context = ssl.create_default_context(cafile=cert_path)
            context_info = f"SSLContext(cafile={cert_path})"
        elif mode == "system":
            if truststore is None:
                return False, "truststore is not installed"
            truststore.inject_into_ssl()
            context = ssl.create_default_context()
            context_info = "system trust store via truststore"
        else:
            context = ssl._create_unverified_context()
            context_info = "unverified SSLContext"
        req = Request(url, headers={"User-Agent": USER_AGENT})
        with urlopen(req, timeout=timeout, context=context) as response:
            response.read(128)
            return True, f"HTTP {response.status}; {context_info}"
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"


def decode_html(content: bytes) -> str:
    for encoding in ("windows-1251", "utf-8"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return content.decode("windows-1251", "replace")


def project_id_from_url(url: str) -> str:
    query = parse_qs(urlparse(url).query)
    project = query.get("proj", [""])[0]
    if not re.fullmatch(r"[A-Fa-f0-9]{32}", project):
        raise ExporterError("The input URL must contain a 32-character proj parameter.")
    return project.upper()


def page_url(bank_url: str, project: str, page: int) -> str:
    return urljoin(bank_url, "questions.php") + "?" + urlencode({"proj": project, "page": page})


def task_url(bank_url: str, project: str, public_id: str) -> str:
    return urljoin(bank_url, "index.php") + "?" + urlencode({"proj": project, "qid": public_id})


def parse_total_count(page_html: str) -> int | None:
    match = re.search(r"setQCount\((\d+)", page_html)
    return int(match.group(1)) if match else None


def parse_page_size(page_html: str) -> int | None:
    match = re.search(r"setQCount\(\s*\d+\s*,\s*\d+\s*,\s*(\d+)", page_html)
    return int(match.group(1)) if match else None


def extract_project_title(index_html: str) -> str | None:
    doc = parse_html(index_html)
    values = [text.strip() for text in doc.xpath("//div[contains(@class,'proj-nav-panel')]//text()") if text.strip()]
    return " ".join(values) if values else None


def parse_html(source: str) -> html.HtmlElement:
    parser = html.HTMLParser(encoding="windows-1251")
    return html.fromstring(source.encode("windows-1251", "xmlcharrefreplace"), parser=parser)


def iter_task_blocks(page_html: str) -> Iterable[tuple[str, html.HtmlElement, html.HtmlElement | None]]:
    doc = parse_html(page_html)
    for block in doc.xpath("//div[contains(concat(' ', normalize-space(@class), ' '), ' qblock ')]"):
        raw_id = block.get("id", "")
        match = re.fullmatch(r"q([A-F0-9]{6})", raw_id)
        if not match:
            continue
        public_id = match.group(1)
        info_nodes = doc.xpath(f"//div[@id='i{public_id}']")
        yield public_id, block, info_nodes[0] if info_nodes else None


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def element_text(node: html.HtmlElement | None) -> str:
    if node is None:
        return ""
    clean_node = copy.deepcopy(node)
    for noisy in clean_node.xpath(".//script|.//style"):
        noisy.drop_tree()
    return normalize_space(clean_node.text_content())


def parse_info(info_node: html.HtmlElement | None) -> dict[str, str | list[str]]:
    metadata: dict[str, str | list[str]] = {"kes": []}
    if info_node is None:
        return metadata
    for row in info_node.xpath(".//tr"):
        cells = row.xpath("./td")
        if len(cells) < 2:
            continue
        name = normalize_space(cells[0].text_content()).rstrip(":")
        value = normalize_space(cells[1].text_content())
        if not name:
            continue
        if name.upper() == "КЭС":
            metadata["kes"] = [normalize_space(v) for v in cells[1].xpath(".//div/text()") if normalize_space(v)]
            if not metadata["kes"] and value:
                metadata["kes"] = [value]
        elif name == "Тип ответа":
            metadata["answer_type"] = value
        else:
            metadata[name.lower()] = value
    return metadata


def extract_guid(block: html.HtmlElement) -> str | None:
    values = block.xpath(".//input[translate(@name,'GUID','guid')='guid']/@value")
    return values[0].upper() if values else None


def extract_options(block: html.HtmlElement) -> list[dict[str, str]]:
    options: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for option in block.xpath(".//option[@value and string-length(normalize-space(@value)) > 0]"):
        value = normalize_space(option.get("value", ""))
        text = normalize_space(option.text_content()) or value
        key = (value, text)
        if value != "0" and key not in seen:
            options.append({"value": value, "text": text})
            seen.add(key)
    for cell in block.xpath(".//*[contains(concat(' ', normalize-space(@class), ' '), ' active-distractor ')]"):
        text = element_text(cell)
        value = cell.get("data-value") or cell.get("value") or text
        key = (value, text)
        if text and key not in seen:
            options.append({"value": value, "text": text})
            seen.add(key)
    return options


def extract_image_refs(fragment: str, block: html.HtmlElement) -> list[str]:
    refs = set(block.xpath(".//img/@src"))
    for call in re.finditer(r"ShowPicture\w*\((.*?)\)", fragment, re.S):
        for ref in re.findall(r"['\"]([^'\"]+)['\"]", call.group(1)):
            if not re.search(r"\.(?:png|jpe?g|gif|svg|webp|bmp|mp4|flv|swf)(?:$|\?)", ref, re.I):
                continue
            if not ref.startswith(("http://", "https://", "../", "/")):
                ref = "../../" + ref
            refs.add(ref)
    for match in re.finditer(r"showXGraphicByGuid\(\s*['\"]([^'\"]+)['\"]", fragment):
        refs.add(f"../../showxgprahic.php?qguid={match.group(1)}")
    return sorted(ref for ref in refs if ref and not ref.startswith("data:"))


def html_fragment(block: html.HtmlElement) -> str:
    return html.tostring(block, encoding="unicode", method="html")


def make_markdown(task: dict) -> str:
    lines = [
        f"# FIPI task {task['qid']}",
        "",
        f"- URL: {task['url']}",
        f"- Project: {task['project']}",
        f"- Internal GUID: {task.get('internal_guid') or ''}",
        f"- Answer type: {task.get('answer_type') or ''}",
    ]
    for kes in task.get("kes", []):
        lines.append(f"- КЭС: {kes}")
    lines.extend(["", task.get("text", ""), ""])
    if task.get("options"):
        lines.append("## Options")
        for option in task["options"]:
            lines.append(f"- {option.get('value')}: {option.get('text')}")
        lines.append("")
    if task.get("local_images"):
        lines.append("## Local images")
        for image_path in task["local_images"]:
            lines.append(f"- {image_path}")
        lines.append("")
    lines.extend(["## Source HTML", "", "```html", task.get("html", ""), "```", ""])
    return "\n".join(lines)


def existing_keys(jsonl_path: Path) -> tuple[set[str], set[str]]:
    ids: set[str] = set()
    urls: set[str] = set()
    if not jsonl_path.exists():
        return ids, urls
    with jsonl_path.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                logging.warning("Ignoring malformed JSONL line %s in %s", line_number, jsonl_path)
                continue
            qid = record.get("qid")
            if qid:
                ids.add(str(qid))
            url = record.get("url")
            if url:
                urls.add(str(url))
    return ids, urls


def read_jsonl(path: Path) -> list[dict]:
    if not path.exists():
        return []
    records = []
    with path.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            if line.strip():
                try:
                    records.append(json.loads(line))
                except json.JSONDecodeError:
                    logging.warning("Ignoring malformed JSONL line %s in %s", line_number, path)
    return records


def append_jsonl(path: Path, record: dict) -> None:
    with path.open("a", encoding="utf-8") as file:
        file.write(json.dumps(record, ensure_ascii=False) + "\n")


def write_csv(path: Path, records: list[dict]) -> None:
    fieldnames = [
        "qid",
        "internal_guid",
        "url",
        "project",
        "project_title",
        "kes",
        "answer_type",
        "text",
        "options",
        "answer",
        "explanation",
        "local_images",
        "source_url",
        "scraped_at",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for record in records:
            row = dict(record)
            for key in ("kes", "options", "local_images"):
                row[key] = json.dumps(row.get(key, []), ensure_ascii=False)
            writer.writerow(row)


def write_sqlite(path: Path, records: list[dict]) -> None:
    conn = sqlite3.connect(path)
    try:
        conn.execute("DROP TABLE IF EXISTS tasks")
        conn.execute(
            """
            CREATE TABLE tasks (
                qid TEXT PRIMARY KEY,
                internal_guid TEXT,
                url TEXT,
                project TEXT,
                project_title TEXT,
                kes TEXT,
                answer_type TEXT,
                text TEXT,
                options TEXT,
                answer TEXT,
                explanation TEXT,
                local_images TEXT,
                html TEXT,
                source_url TEXT,
                scraped_at TEXT
            )
            """
        )
        for record in records:
            conn.execute(
                """
                INSERT OR REPLACE INTO tasks VALUES (
                    :qid, :internal_guid, :url, :project, :project_title, :kes, :answer_type, :text,
                    :options, :answer, :explanation, :local_images, :html, :source_url, :scraped_at
                )
                """,
                {
                    **record,
                    "kes": json.dumps(record.get("kes", []), ensure_ascii=False),
                    "options": json.dumps(record.get("options", []), ensure_ascii=False),
                    "local_images": json.dumps(record.get("local_images", []), ensure_ascii=False),
                },
            )
        conn.commit()
    finally:
        conn.close()


def download_images(
    client: FipiClient,
    image_refs: list[str],
    task_page_url: str,
    out_dir: Path,
    project: str,
) -> tuple[list[str], int, int]:
    local_paths: list[str] = []
    downloaded = 0
    errors = 0
    out_dir.mkdir(parents=True, exist_ok=True)
    for ref in image_refs:
        image_url = urljoin(task_page_url, ref)
        parsed = urlparse(image_url)
        suffix = Path(parsed.path).suffix.lower() or ".bin"
        digest = hashlib.sha1(image_url.encode("utf-8")).hexdigest()[:12]
        file_name = f"{project}_{digest}{suffix}"
        target = out_dir / file_name
        if not target.exists():
            try:
                result = client.fetch(image_url, referer=task_page_url, binary=True)
            except ExporterError as exc:
                logging.warning("Skipping image %s: %s", image_url, exc)
                errors += 1
                continue
            target.write_bytes(result.content)
            downloaded += 1
        local_paths.append(str(target))
    return local_paths, downloaded, errors


def parse_task(
    public_id: str,
    block: html.HtmlElement,
    info_node: html.HtmlElement | None,
    page_source: str,
    task_page_url: str,
    project: str,
    project_title: str | None,
    local_images: list[str],
    source_url: str,
) -> dict:
    fragment = html_fragment(block)
    info = parse_info(info_node)
    text = element_text(block)
    return {
        "qid": public_id,
        "internal_guid": extract_guid(block),
        "url": task_url(task_page_url, project, public_id),
        "project": project,
        "project_title": project_title,
        "kes": info.get("kes", []),
        "theme": None,
        "task_number": public_id,
        "answer_type": info.get("answer_type"),
        "text": text,
        "options": extract_options(block),
        "answer": None,
        "explanation": None,
        "image_refs": extract_image_refs(fragment, block),
        "local_images": local_images,
        "html": fragment,
        "scraped_at": datetime.now(timezone.utc).isoformat(),
        "source_url": source_url,
        "page_html_sha1": hashlib.sha1(page_source.encode("utf-8", "replace")).hexdigest(),
    }


def export(args: argparse.Namespace) -> ExportStats:
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    images_dir = out_dir / "images"
    markdown_dir = out_dir / "markdown"
    markdown_dir.mkdir(parents=True, exist_ok=True)
    debug_dir = out_dir / "debug" if args.debug else None
    jsonl_path = out_dir / "tasks.jsonl"
    csv_path = out_dir / "tasks.csv"
    sqlite_path = out_dir / "tasks.sqlite"

    project = project_id_from_url(args.url)
    bank_url = urljoin(args.url, ".")
    client = FipiClient(args.delay, args.timeout, debug_dir, tls_mode=args.tls_mode)
    client.log_configuration()
    stats = ExportStats()

    logging.info("Fetching index page")
    index_result = client.fetch(args.url)
    index_html = decode_html(index_result.content)
    project_title = extract_project_title(index_html)

    seen_ids, seen_urls = existing_keys(jsonl_path) if args.resume else (set(), set())
    if jsonl_path.exists() and not args.resume:
        jsonl_path.unlink()
        seen_ids, seen_urls = set(), set()

    total_count: int | None = None
    page_size: int | None = None
    max_page: int | None = None
    saved_this_run = 0
    page = 0

    while True:
        if args.limit and saved_this_run >= args.limit:
            break
        if max_page is not None and page > max_page:
            break
        current_url = page_url(bank_url, project, page)
        logging.info("Fetching page %s", page)
        result = client.fetch(current_url, referer=args.url)
        stats.pages_fetched += 1
        page_html = decode_html(result.content)
        total_count = total_count or parse_total_count(page_html)
        page_size = page_size or parse_page_size(page_html)
        if total_count and page_size:
            max_page = max(0, math.ceil(total_count / page_size) - 1)
        blocks = list(iter_task_blocks(page_html))
        if not blocks:
            logging.info("No task blocks found on page %s; stopping", page)
            break

        iterator = blocks
        if tqdm and sys.stderr.isatty():
            iterator = tqdm(blocks, desc=f"page {page}", leave=False)
        for public_id, block, info_node in iterator:
            stats.found += 1
            current_task_url = task_url(current_url, project, public_id)
            if public_id in seen_ids or current_task_url in seen_urls:
                stats.skipped += 1
                stats.duplicates += 1
                continue
            try:
                fragment = html_fragment(block)
                image_refs = extract_image_refs(fragment, block)
                local_images, downloaded, image_errors = download_images(client, image_refs, current_url, images_dir, project)
                stats.images_downloaded += downloaded
                stats.image_errors += image_errors
                task = parse_task(
                    public_id,
                    block,
                    info_node,
                    page_html,
                    current_url,
                    project,
                    project_title,
                    local_images,
                    args.url,
                )
                if not task["text"]:
                    stats.empty_text += 1
                    logging.warning("Task %s has empty text", public_id)
                append_jsonl(jsonl_path, task)
                (markdown_dir / f"{public_id}.md").write_text(make_markdown(task), encoding="utf-8")
                seen_ids.add(public_id)
                seen_urls.add(task["url"])
                stats.saved += 1
                saved_this_run += 1
            except Exception as exc:
                stats.task_errors += 1
                logging.exception("Skipping task %s after parsing/export error: %s", public_id, exc)
            if args.limit and saved_this_run >= args.limit:
                break

        page += 1

    records = read_jsonl(jsonl_path)
    formats = {part.strip().lower() for part in args.format.split(",") if part.strip()}
    if "csv" in formats:
        write_csv(csv_path, records)
    if "sqlite" in formats:
        write_sqlite(sqlite_path, records)
    if "jsonl" not in formats:
        logging.warning("JSONL is always written for resume support: %s", jsonl_path)

    logging.info("Total tasks reported by site: %s", total_count or "unknown")
    logging.info("Tasks found this run: %s", stats.found)
    logging.info("Tasks saved this run: %s", stats.saved)
    logging.info("Tasks skipped this run: %s", stats.skipped)
    logging.info("Duplicates skipped this run: %s", stats.duplicates)
    logging.info("Tasks with empty text: %s", stats.empty_text)
    logging.info("Images downloaded this run: %s", stats.images_downloaded)
    logging.info("Image download errors this run: %s", stats.image_errors)
    logging.info("Task export errors this run: %s", stats.task_errors)
    logging.info("Output directory: %s", out_dir.resolve())
    return stats


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Export tasks from the public FIPI OGE task bank.")
    parser.add_argument("--url", default=DEFAULT_URL, help="FIPI bank URL with proj parameter.")
    parser.add_argument("--out", default="data", help="Output directory.")
    parser.add_argument("--delay", type=float, default=2.0, help="Delay between requests in seconds.")
    parser.add_argument("--timeout", type=float, default=30.0, help="HTTP timeout in seconds.")
    parser.add_argument("--limit", type=int, default=None, help="Maximum new tasks to save.")
    parser.add_argument("--resume", action="store_true", help="Skip qids already present in tasks.jsonl.")
    parser.add_argument("--format", default="jsonl,csv,sqlite", help="Comma-separated export formats.")
    parser.add_argument("--headless", default="true", help="Accepted for compatibility; Playwright is not used.")
    parser.add_argument("--debug", action="store_true", help="Save fetched HTML pages under data/debug.")
    parser.add_argument("--diagnose-tls", action="store_true", help="Run TLS diagnostics for --url and exit.")
    parser.add_argument(
        "--tls-mode",
        choices=("certifi", "system", "insecure"),
        default="certifi",
        help="TLS trust mode: certifi bundle, Windows/system trust store via truststore, or explicit insecure mode.",
    )
    parser.add_argument(
        "--insecure",
        action="store_true",
        help="Deprecated alias for --tls-mode insecure. Use only for local certificate-chain diagnostics.",
    )
    return parser


def configure_logging(debug: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    configure_logging(args.debug)
    if args.delay < 1:
        logging.warning("A delay below 1 second is not recommended for this public site.")
    if args.insecure:
        args.tls_mode = "insecure"
    if args.diagnose_tls:
        return diagnose_tls(args.url, args.timeout)
    try:
        export(args)
    except ExporterError as exc:
        logging.error("%s", exc)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
