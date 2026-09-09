
from __future__ import annotations

import argparse
import csv
import io
import json
import logging
import os
import random
import re
import socket
import time
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
import google_auth_httplib2
import httplib2
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CSV_MIME = "text/csv"


def extract_file_id(value: str) -> str:
    value = value.strip()
    match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", value)
    if match:
        return match.group(1)

    if re.fullmatch(r"[A-Za-z0-9_-]{10,}", value):
        return value

    raise ValueError("无法识别 Google Sheet URL / file_id。")


def safe_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().rstrip(".")
    return name or "google_sheet"


def load_credentials(credentials_path: Path, token_path: Path) -> Credentials:
    creds = None

    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
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
                "请在 Google Cloud Console 创建 Desktop OAuth Client，"
                "并下载为 credentials.json。"
            )

        flow = InstalledAppFlow.from_client_secrets_file(
            str(credentials_path),
            SCOPES,
        )
        creds = flow.run_local_server(port=0)

        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return creds


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        logging.exception("状态文件损坏，将按首次运行处理。")
        return {}


def save_state_atomic(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(tmp, path)


def call_with_retry(func, attempts: int = 5):
    """Retry transient Google API and network failures with exponential backoff."""
    delay = 1.0

    for attempt in range(1, attempts + 1):
        try:
            return func()

        except HttpError as exc:
            status = getattr(exc.resp, "status", None)
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


def configure_proxy_environment(proxy_url: str | None) -> None:
    """
    Make OAuth token refresh (requests) and other HTTP clients use the same proxy.
    """
    if not proxy_url:
        return

    os.environ["HTTP_PROXY"] = proxy_url
    os.environ["HTTPS_PROXY"] = proxy_url
    os.environ["http_proxy"] = proxy_url
    os.environ["https_proxy"] = proxy_url


def build_google_service(api_name: str, api_version: str, creds, proxy_url: str | None):
    """
    Build a Google API service with an explicit httplib2 proxy.

    Passing an AuthorizedHttp explicitly avoids relying on VS Code / Windows
    system-proxy inheritance.
    """
    if proxy_url:
        proxy_info = httplib2.proxy_info_from_url(proxy_url, method="https")
        raw_http = httplib2.Http(
            proxy_info=proxy_info,
            timeout=30,
        )
    else:
        raw_http = httplib2.Http(timeout=30)

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


def get_metadata(drive, file_id: str) -> dict[str, Any]:
    return call_with_retry(
        lambda: drive.files().get(
            fileId=file_id,
            fields="id,name,mimeType,version,modifiedTime",
        ).execute()
    )


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
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = call_with_retry(downloader.next_chunk)

        if tmp_path.stat().st_size == 0:
            raise RuntimeError(f"导出结果为空：{output_path.name}")

        os.replace(tmp_path, output_path)

    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def get_sheet_tabs(sheets, spreadsheet_id: str) -> list[dict[str, Any]]:
    result = call_with_retry(
        lambda: sheets.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId,title,index))",
        ).execute()
    )

    tabs = []
    for item in result.get("sheets", []):
        props = item["properties"]
        tabs.append(
            {
                "sheet_id": props["sheetId"],
                "title": props["title"],
                "index": props["index"],
            }
        )

    return sorted(tabs, key=lambda x: x["index"])


def export_each_sheet_as_csv(
    sheets,
    *,
    spreadsheet_id: str,
    spreadsheet_name: str,
    output_dir: Path,
) -> list[Path]:
    """
    CSV 通过 Sheets API values.get 导出。
    这样可以稳定地对每个 tab 单独生成 CSV，
    而不是依赖网页 export URL / gid 参数。
    """
    tabs = get_sheet_tabs(sheets, spreadsheet_id)
    output_paths: list[Path] = []

    for tab in tabs:
        title = tab["title"]

        response = call_with_retry(
            lambda title=title: sheets.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title.replace(chr(39), chr(39) * 2)}'",
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
        tmp_path = output_path.with_suffix(".csv.tmp")
        output_dir.mkdir(parents=True, exist_ok=True)

        try:
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


def sync_once(
    drive,
    sheets,
    *,
    file_id: str,
    output_dir: Path,
    state_path: Path,
    force: bool = False,
) -> bool:
    metadata = get_metadata(drive, file_id)

    if metadata.get("mimeType") != SHEET_MIME:
        raise ValueError("目标文件不是 Google Sheets 文件。")

    name = metadata["name"]
    version = str(metadata.get("version", ""))
    modified_time = metadata.get("modifiedTime")

    state = load_state(state_path)
    previous_version = str(state.get("version", ""))

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

    # 1. 整个 Spreadsheet -> 一个 XLSX
    xlsx_path = output_dir / f"{safe_filename(name)}.xlsx"
    download_export_atomic(
        drive,
        file_id=file_id,
        mime_type=XLSX_MIME,
        output_path=xlsx_path,
    )

    # 2. 每个工作表 -> 一个 CSV
    csv_paths = export_each_sheet_as_csv(
        sheets,
        spreadsheet_id=file_id,
        spreadsheet_name=name,
        output_dir=output_dir,
    )

    # 两种格式都成功后，才推进本地 version。
    # 如果中途失败，下轮仍会重新同步，不会误认为已经完成。
    save_state_atomic(
        state_path,
        {
            "file_id": file_id,
            "name": name,
            "version": version,
            "modified_time": modified_time,
            "xlsx": str(xlsx_path),
            "csv": [str(p) for p in csv_paths],
        },
    )

    logging.info("XLSX 已更新：%s", xlsx_path)
    for path in csv_paths:
        logging.info("CSV 已更新：%s", path)

    return True


def main() -> int:
    parser = argparse.ArgumentParser(
        description="监控 Google Sheet 更新，并自动导出 XLSX + 每工作表 CSV。"
    )
    parser.add_argument(
        "sheet",
        help="Google Sheets URL 或 file_id",
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
        help="下载目录，默认 ./downloads",
    )
    parser.add_argument(
        "--credentials",
        default="credentials.json",
        help="Google OAuth Desktop Client JSON",
    )
    parser.add_argument(
        "--token",
        default="token.json",
        help="授权 token 保存路径",
    )
    parser.add_argument(
        "--proxy",
        default="http://127.0.0.1:7890",
        help=(
            "HTTP/HTTPS 代理地址，默认 http://127.0.0.1:7890；"
            "例如 Clash/Mihomo Mixed Port。传空字符串可关闭代理。"
        ),
    )
    parser.add_argument(
        "--state",
        default=".sheet_watcher_state.json",
        help="同步状态文件",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="只检查并同步一次",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="忽略版本状态，强制导出一次",
    )
    args = parser.parse_args()

    if args.interval < 10:
        parser.error("--interval 最低建议 10 秒；不要过度轮询 Google API。")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )

    file_id = extract_file_id(args.sheet)
    credentials_path = Path(args.credentials).resolve()
    token_path = Path(args.token).resolve()
    output_dir = Path(args.output).resolve()
    state_path = Path(args.state).resolve()

    proxy_url = args.proxy.strip() or None
    configure_proxy_environment(proxy_url)

    if proxy_url:
        logging.info("Google API 使用代理：%s", proxy_url)
    else:
        logging.info("Google API 未使用显式代理。")

    creds = load_credentials(credentials_path, token_path)

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

    while True:
        try:
            sync_once(
                drive,
                sheets,
                file_id=file_id,
                output_dir=output_dir,
                state_path=state_path,
                force=args.force,
            )
        except KeyboardInterrupt:
            logging.info("收到退出指令。")
            return 0
        except Exception:
            logging.exception("本轮同步失败。下轮会继续重试。")

        if args.once:
            return 0

        # force 只对首次循环生效，避免无限重复下载。
        args.force = False

        try:
            time.sleep(args.interval)
        except KeyboardInterrupt:
            logging.info("收到退出指令。")
            return 0


if __name__ == "__main__":
    raise SystemExit(main())
