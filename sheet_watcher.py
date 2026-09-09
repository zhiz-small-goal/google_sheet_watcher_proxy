from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import random
import re
import socket
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import google_auth_httplib2
import httplib2
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

DEFAULT_PROXY = "http://127.0.0.1:7890"
STATE_SCHEMA_VERSION = 2


# ----------------------------
# Basic helpers
# ----------------------------

def extract_file_id(value: str) -> str:
    """Accept a Google Sheets URL or a raw Drive file ID."""
    value = value.strip()

    match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", value)
    if match:
        return match.group(1)

    if re.fullmatch(r"[A-Za-z0-9_-]{10,}", value):
        return value

    raise ValueError(f"无法识别 Google Sheet URL / file_id：{value!r}")


def safe_filename(name: str) -> str:
    """Make a Windows-safe file/folder name."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().rstrip(". ")
    return name or "google_sheet"


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def quote_sheet_title(title: str) -> str:
    """A1 notation sheet-title quoting."""
    return "'" + title.replace("'", "''") + "'"


# ----------------------------
# OAuth + proxy
# ----------------------------

def configure_proxy_environment(proxy_url: str | None) -> None:
    """
    requests is used by OAuth token refresh.
    Explicitly set proxy env vars so refresh uses the same proxy.
    """
    keys = ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy")

    if proxy_url:
        for key in keys:
            os.environ[key] = proxy_url
    else:
        for key in keys:
            os.environ.pop(key, None)


def make_proxy_info(proxy_url: str | None):
    if not proxy_url:
        return None

    parsed = urlparse(proxy_url)

    if parsed.scheme.lower() not in {"http", "https"}:
        raise ValueError(
            "--proxy 当前支持 http:// 或 https:// 代理，例如 "
            "http://127.0.0.1:7890"
        )

    if not parsed.hostname or not parsed.port:
        raise ValueError(
            f"代理地址缺少 host/port：{proxy_url!r}，"
            "例如 http://127.0.0.1:7890"
        )

    return httplib2.ProxyInfo(
        proxy_type=httplib2.socks.PROXY_TYPE_HTTP,
        proxy_host=parsed.hostname,
        proxy_port=parsed.port,
        proxy_user=parsed.username,
        proxy_pass=parsed.password,
    )


def load_credentials(
    credentials_path: Path,
    token_path: Path,
) -> Credentials:
    creds = None

    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(
                str(token_path),
                SCOPES,
            )
        except Exception:
            logging.exception("token.json 无法读取，将重新授权。")

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            logging.exception("刷新授权失败，将重新授权。")
            creds = None

    if not creds or not creds.valid:
        if not credentials_path.exists():
            raise FileNotFoundError(
                f"缺少 {credentials_path}。\n"
                "请创建 Google OAuth Desktop Client，"
                "并把下载的 JSON 保存为 credentials.json。"
            )

        flow = InstalledAppFlow.from_client_secrets_file(
            str(credentials_path),
            SCOPES,
        )
        creds = flow.run_local_server(port=0)

        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return creds


def build_google_service(
    api_name: str,
    api_version: str,
    creds: Credentials,
    proxy_url: str | None,
):
    proxy_info = make_proxy_info(proxy_url)

    raw_http = httplib2.Http(
        proxy_info=proxy_info,
        timeout=30,
    )

    authorized_http = google_auth_httplib2.AuthorizedHttp(
        creds,
        http=raw_http,
    )

    return build(
        api_name,
        api_version,
        http=authorized_http,
        cache_discovery=False,
    )


# ----------------------------
# Retry
# ----------------------------

def call_with_retry(func, attempts: int = 5):
    """
    Retry transient Google API/network failures with exponential backoff.
    """
    delay = 1.0

    for attempt in range(1, attempts + 1):
        try:
            return func()

        except HttpError as exc:
            status = getattr(exc.resp, "status", None)

            # Permission / auth / bad request etc. are not transient.
            if status not in {408, 429, 500, 502, 503, 504}:
                raise

            error = exc

        except (
            TimeoutError,
            socket.timeout,
            ConnectionError,
            httplib2.HttpLib2Error,
            OSError,
        ) as exc:
            error = exc

        if attempt == attempts:
            raise error

        sleep_for = delay + random.uniform(0, 0.5)

        logging.warning(
            "网络/API 请求失败：%s；%.1f 秒后重试 (%d/%d)",
            error,
            sleep_for,
            attempt,
            attempts,
        )

        time.sleep(sleep_for)
        delay = min(delay * 2, 30)


# ----------------------------
# State
# ----------------------------

def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {
            "schema_version": STATE_SCHEMA_VERSION,
            "files": {},
        }

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        logging.exception("状态文件损坏，将按首次运行处理。")
        return {
            "schema_version": STATE_SCHEMA_VERSION,
            "files": {},
        }

    # Migration from v1 single-file state:
    # {
    #   "file_id": "...",
    #   "name": "...",
    #   "version": "..."
    # }
    if "files" not in raw and raw.get("file_id"):
        file_id = str(raw["file_id"])
        raw = {
            "schema_version": STATE_SCHEMA_VERSION,
            "files": {
                file_id: {
                    "name": raw.get("name"),
                    "version": str(raw.get("version", "")),
                    "modified_time": raw.get("modified_time"),
                    "xlsx": raw.get("xlsx"),
                    "csv": raw.get("csv", []),
                    "migrated_from_v1": True,
                }
            },
        }

    raw.setdefault("schema_version", STATE_SCHEMA_VERSION)
    raw.setdefault("files", {})
    return raw


def save_state_atomic(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(tmp, path)


# ----------------------------
# Drive discovery
# ----------------------------

def get_metadata(drive, file_id: str) -> dict[str, Any]:
    return call_with_retry(
        lambda: drive.files().get(
            fileId=file_id,
            fields=(
                "id,name,mimeType,version,modifiedTime,"
                "ownedByMe,capabilities(canDownload)"
            ),
            supportsAllDrives=True,
        ).execute()
    )


def list_all_spreadsheets(
    drive,
    *,
    owned_only: bool = False,
) -> list[dict[str, Any]]:
    """
    Discover Google Sheets visible to the authorized Drive account.

    --all:
        all visible, non-trashed Google Sheets
    --all --owned-only:
        only Google Sheets owned by this account
    """
    q_parts = [
        f"mimeType='{SHEET_MIME}'",
        "trashed=false",
    ]

    if owned_only:
        q_parts.append("'me' in owners")

    query = " and ".join(q_parts)

    files: list[dict[str, Any]] = []
    page_token = None

    while True:
        result = call_with_retry(
            lambda page_token=page_token: drive.files().list(
                q=query,
                spaces="drive",
                corpora="user",
                pageSize=1000,
                pageToken=page_token,
                fields=(
                    "nextPageToken,"
                    "files(id,name,mimeType,version,modifiedTime,"
                    "ownedByMe,capabilities(canDownload))"
                ),
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
            ).execute()
        )

        files.extend(result.get("files", []))
        page_token = result.get("nextPageToken")

        if not page_token:
            break

    files.sort(key=lambda item: (item.get("name", "").casefold(), item["id"]))
    return files


# ----------------------------
# Watchlist
# ----------------------------

def load_watchlist(path: Path) -> list[str]:
    """
    Supported JSON:
      ["URL", "file_id"]
    or
      {"sheets": ["URL", {"file_id": "..."}, {"url": "..."}]}

    Supported TXT:
      one URL / file ID per line; # comments allowed.
    """
    if not path.exists():
        raise FileNotFoundError(f"watchlist 不存在：{path}")

    values: list[str] = []

    if path.suffix.lower() == ".json":
        data = json.loads(path.read_text(encoding="utf-8"))

        if isinstance(data, dict):
            data = data.get("sheets", [])

        if not isinstance(data, list):
            raise ValueError("watchlist JSON 必须是数组，或包含 sheets 数组。")

        for item in data:
            if isinstance(item, str):
                values.append(item)
            elif isinstance(item, dict):
                value = item.get("file_id") or item.get("url")
                if value:
                    values.append(str(value))
                else:
                    raise ValueError(
                        f"watchlist 条目缺少 file_id/url：{item!r}"
                    )
            else:
                raise ValueError(f"不支持的 watchlist 条目：{item!r}")

    else:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            values.append(line)

    file_ids: list[str] = []
    seen: set[str] = set()

    for value in values:
        file_id = extract_file_id(value)
        if file_id not in seen:
            seen.add(file_id)
            file_ids.append(file_id)

    if not file_ids:
        raise ValueError("watchlist 为空。")

    return file_ids


def resolve_watchlist_metadata(
    drive,
    watchlist_path: Path,
) -> list[dict[str, Any]]:
    metadata: list[dict[str, Any]] = []

    for file_id in load_watchlist(watchlist_path):
        item = get_metadata(drive, file_id)

        if item.get("mimeType") != SHEET_MIME:
            raise ValueError(
                f"{item.get('name', file_id)} 不是 Google Sheets 文件。"
            )

        metadata.append(item)

    metadata.sort(key=lambda item: (item.get("name", "").casefold(), item["id"]))
    return metadata


# ----------------------------
# Export
# ----------------------------

def make_output_dir(
    root: Path,
    metadata: dict[str, Any],
) -> Path:
    """
    Include short file_id to prevent collisions between same-named Sheets.
    """
    name = safe_filename(metadata["name"])
    short_id = metadata["id"][:8]
    return root / f"{name}__{short_id}"


def download_export_atomic(
    drive,
    *,
    file_id: str,
    mime_type: str,
    output_path: Path,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")

    try:
        request = drive.files().export_media(
            fileId=file_id,
            mimeType=mime_type,
        )

        with tmp_path.open("wb") as fh:
            downloader = MediaIoBaseDownload(
                fh,
                request,
                chunksize=1024 * 1024,
            )

            done = False
            while not done:
                _, done = call_with_retry(downloader.next_chunk)

        if not tmp_path.exists() or tmp_path.stat().st_size == 0:
            raise RuntimeError(f"导出结果为空：{output_path.name}")

        os.replace(tmp_path, output_path)

    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def get_sheet_tabs(
    sheets,
    spreadsheet_id: str,
) -> list[dict[str, Any]]:
    result = call_with_retry(
        lambda: sheets.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId,title,index,hidden))",
        ).execute()
    )

    tabs: list[dict[str, Any]] = []

    for item in result.get("sheets", []):
        props = item["properties"]

        tabs.append(
            {
                "sheet_id": props["sheetId"],
                "title": props["title"],
                "index": props["index"],
                "hidden": bool(props.get("hidden", False)),
            }
        )

    return sorted(tabs, key=lambda item: item["index"])


def export_each_sheet_as_csv(
    sheets,
    *,
    spreadsheet_id: str,
    spreadsheet_name: str,
    output_dir: Path,
) -> list[Path]:
    """
    Export every tab, including hidden tabs, as a separate CSV.

    CSV contains formatted/displayed values, not formulas.
    """
    tabs = get_sheet_tabs(sheets, spreadsheet_id)
    output_paths: list[Path] = []

    for tab in tabs:
        title = tab["title"]

        response = call_with_retry(
            lambda title=title: sheets.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=quote_sheet_title(title),
                valueRenderOption="FORMATTED_VALUE",
                dateTimeRenderOption="FORMATTED_STRING",
            ).execute()
        )

        values = response.get("values", [])

        filename = (
            f"{safe_filename(spreadsheet_name)}"
            f"__{safe_filename(title)}.csv"
        )
        output_path = output_dir / filename
        tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")

        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            # utf-8-sig makes Chinese CSV friendlier in Windows Excel.
            with tmp_path.open(
                "w",
                encoding="utf-8-sig",
                newline="",
            ) as fh:
                writer = csv.writer(fh, lineterminator="\n")
                writer.writerows(values)

            os.replace(tmp_path, output_path)

        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass

        output_paths.append(output_path)

    return output_paths


# ----------------------------
# Sync
# ----------------------------

def sync_one(
    drive,
    sheets,
    *,
    metadata: dict[str, Any],
    output_root: Path,
    state: dict[str, Any],
    state_path: Path,
    force: bool = False,
) -> bool:
    file_id = metadata["id"]
    name = metadata["name"]
    version = str(metadata.get("version", ""))
    modified_time = metadata.get("modifiedTime")

    if metadata.get("mimeType") != SHEET_MIME:
        raise ValueError(f"{name} 不是 Google Sheets 文件。")

    can_download = metadata.get("capabilities", {}).get("canDownload", True)
    if not can_download:
        raise PermissionError(f"{name} 当前权限不允许下载/导出。")

    previous = state["files"].get(file_id, {})
    previous_version = str(previous.get("version", ""))

    if not force and previous_version == version:
        logging.info(
            "无更新：%s | version=%s",
            name,
            version,
        )
        return False

    logging.info(
        "检测到更新：%s | %s -> %s | modified=%s",
        name,
        previous_version or "(首次)",
        version,
        modified_time,
    )

    output_dir = make_output_dir(output_root, metadata)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1) Whole Spreadsheet -> XLSX
    xlsx_path = output_dir / f"{safe_filename(name)}.xlsx"

    download_export_atomic(
        drive,
        file_id=file_id,
        mime_type=XLSX_MIME,
        output_path=xlsx_path,
    )

    # 2) Every tab -> one CSV
    csv_paths = export_each_sheet_as_csv(
        sheets,
        spreadsheet_id=file_id,
        spreadsheet_name=name,
        output_dir=output_dir,
    )

    # Only advance version after ALL exports succeed.
    state["schema_version"] = STATE_SCHEMA_VERSION
    state["files"][file_id] = {
        "name": name,
        "version": version,
        "modified_time": modified_time,
        "owned_by_me": metadata.get("ownedByMe"),
        "output_dir": str(output_dir),
        "xlsx": str(xlsx_path),
        "csv": [str(path) for path in csv_paths],
        "last_success_utc": utc_now_iso(),
    }

    save_state_atomic(state_path, state)

    logging.info("同步完成：%s", name)
    logging.info("  XLSX：%s", xlsx_path)
    logging.info("  CSV：%d 个", len(csv_paths))

    return True


def sync_targets(
    drive,
    sheets,
    *,
    targets: list[dict[str, Any]],
    output_root: Path,
    state_path: Path,
    force: bool,
) -> tuple[int, int, int]:
    """
    Returns:
      (updated_count, unchanged_count, failed_count)
    """
    state = load_state(state_path)

    updated = 0
    unchanged = 0
    failed = 0

    for metadata in targets:
        try:
            changed = sync_one(
                drive,
                sheets,
                metadata=metadata,
                output_root=output_root,
                state=state,
                state_path=state_path,
                force=force,
            )

            if changed:
                updated += 1
            else:
                unchanged += 1

        except KeyboardInterrupt:
            raise

        except Exception:
            failed += 1
            logging.exception(
                "同步失败：%s (%s)。其他文件继续处理；下轮会重试。",
                metadata.get("name", "(unknown)"),
                metadata.get("id", "(unknown)"),
            )

    return updated, unchanged, failed


# ----------------------------
# CLI
# ----------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "监控一个、指定列表或账号下全部 Google Sheets，"
            "更新后自动导出 XLSX + 每个 tab 的 CSV。"
        )
    )

    parser.add_argument(
        "sheet",
        nargs="?",
        help="单文件模式：Google Sheets URL 或 file_id",
    )

    mode = parser.add_mutually_exclusive_group()

    mode.add_argument(
        "--all",
        action="store_true",
        help="自动发现并监控账号当前可见的全部 Google Sheets",
    )

    mode.add_argument(
        "--watchlist",
        help="只监控 watchlist JSON/TXT 中指定的 Google Sheets",
    )

    parser.add_argument(
        "--owned-only",
        action="store_true",
        help="配合 --all：只监控当前账号拥有的 Sheets，忽略别人共享的文件",
    )

    parser.add_argument(
        "--interval",
        type=int,
        default=60,
        help="轮询间隔秒数，默认 60",
    )

    parser.add_argument(
        "--output",
        default="downloads",
        help="下载根目录，默认 ./downloads",
    )

    parser.add_argument(
        "--credentials",
        default="credentials.json",
        help="Google OAuth Desktop Client JSON",
    )

    parser.add_argument(
        "--token",
        default="token.json",
        help="OAuth token 保存路径",
    )

    parser.add_argument(
        "--state",
        default=".sheet_watcher_state.json",
        help="同步状态文件",
    )

    parser.add_argument(
        "--proxy",
        default=DEFAULT_PROXY,
        help=(
            f"HTTP/HTTPS 代理，默认 {DEFAULT_PROXY}；"
            '传 --proxy "" 可关闭'
        ),
    )

    parser.add_argument(
        "--once",
        action="store_true",
        help="只扫描/同步一次后退出",
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help="忽略已记录版本，强制重新导出",
    )

    args = parser.parse_args()

    if args.interval < 10:
        parser.error("--interval 最低 10 秒，避免过度轮询 API。")

    selected_modes = int(bool(args.sheet)) + int(args.all) + int(bool(args.watchlist))

    if selected_modes != 1:
        parser.error(
            "必须且只能选择一种模式："
            "提供单个 Sheet URL/file_id，或 --all，或 --watchlist。"
        )

    if args.owned_only and not args.all:
        parser.error("--owned-only 只能和 --all 一起使用。")

    return args


def resolve_targets(drive, args) -> list[dict[str, Any]]:
    if args.sheet:
        file_id = extract_file_id(args.sheet)
        metadata = get_metadata(drive, file_id)

        if metadata.get("mimeType") != SHEET_MIME:
            raise ValueError("目标文件不是 Google Sheets 文件。")

        return [metadata]

    if args.all:
        return list_all_spreadsheets(
            drive,
            owned_only=args.owned_only,
        )

    return resolve_watchlist_metadata(
        drive,
        Path(args.watchlist).resolve(),
    )


def main() -> int:
    args = parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )

    proxy_url = args.proxy.strip() or None
    configure_proxy_environment(proxy_url)

    if proxy_url:
        logging.info("Google API 使用代理：%s", proxy_url)
    else:
        logging.info("Google API 未使用显式代理。")

    credentials_path = Path(args.credentials).resolve()
    token_path = Path(args.token).resolve()
    output_root = Path(args.output).resolve()
    state_path = Path(args.state).resolve()

    creds = load_credentials(
        credentials_path,
        token_path,
    )

    drive = build_google_service(
        "drive",
        "v3",
        creds,
        proxy_url,
    )

    sheets = build_google_service(
        "sheets",
        "v4",
        creds,
        proxy_url,
    )

    first_loop = True

    while True:
        try:
            targets = resolve_targets(drive, args)

            if not targets:
                logging.warning("本轮没有找到任何 Google Sheets。")
            else:
                logging.info(
                    "本轮目标：%d 个 Google Sheets",
                    len(targets),
                )

                for item in targets:
                    logging.info(
                        "  - %s | %s",
                        item.get("name", "(unknown)"),
                        item.get("id", "(unknown)"),
                    )

                updated, unchanged, failed = sync_targets(
                    drive,
                    sheets,
                    targets=targets,
                    output_root=output_root,
                    state_path=state_path,
                    force=(args.force and first_loop),
                )

                logging.info(
                    "本轮完成：更新=%d，无变化=%d，失败=%d",
                    updated,
                    unchanged,
                    failed,
                )

        except KeyboardInterrupt:
            logging.info("收到退出指令。")
            return 0

        except Exception:
            logging.exception("本轮扫描失败。下轮会继续重试。")

        if args.once:
            return 0

        first_loop = False

        try:
            time.sleep(args.interval)
        except KeyboardInterrupt:
            logging.info("收到退出指令。")
            return 0


if __name__ == "__main__":
    raise SystemExit(main())
