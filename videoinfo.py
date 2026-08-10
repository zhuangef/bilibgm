#!/usr/bin/env python3
"""Pythonista 专用:批量将 B 站视频 TAG 名称和 music_id 写入 Excel 文件。

脚本会从同目录 ``video_list.xlsx`` 的 A 列读取 BVID,先查询每个视频的 CID,
再调用支持 BGM 信息的 B 站 TAG 接口。TAG 名称会写入原工作簿的 R 列,
music_id 会写入 S 列。

Pythonista 用法:
    1. 把本脚本和 video_list.xlsx 放在同一个目录。
    2. 如需登录态,把 Cookie 填到下方 PYTHONISTA_COOKIE 常量。
    3. 在 Pythonista 中直接运行本脚本。

依赖:
    Pythonista 需要先安装 requests 和 openpyxl。可在 StaSh 中执行:
    pip install requests openpyxl
"""

from __future__ import annotations

import re
import shutil
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any


VIEW_URL = "https://api.bilibili.com/x/web-interface/view"
TAG_URL = "https://api.bilibili.com/x/web-interface/view/detail/tag"

BVID_COLUMN = "A"
TAG_NAME_COLUMN = "R"
MUSIC_ID_COLUMN = "S"
BOOK_TITLE_PATTERN = re.compile(r"《([^》]+)》")

PYTHONISTA_EXCEL_FILENAME = "video_list.xlsx"
# 在 Pythonista 中运行时,请把 Cookie 写在这里,例如:
# PYTHONISTA_COOKIE = "SESSDATA=xxx; bili_jct=xxx"
PYTHONISTA_COOKIE = ""
PYTHONISTA_START_ROW = 2
PYTHONISTA_SEPARATOR = "; "
PYTHONISTA_TIMEOUT = 10.0
PYTHONISTA_SLEEP = 0.3
PYTHONISTA_CREATE_BACKUP = True


@dataclass
class Options:
    """Pythonista 脚本运行参数。"""

    excel: Path
    start_row: int = 2
    separator: str = "; "
    cookie: str = ""
    timeout: float = 10.0
    sleep: float = 0.3
    no_backup: bool = False


class BilibiliAPIError(RuntimeError):
    """B 站接口返回错误时抛出。"""


def get_script_dir() -> Path:
    """返回脚本所在目录,兼容没有 __file__ 的运行方式。"""
    script = globals().get("__file__")
    if script:
        return Path(script).resolve().parent
    return Path.cwd()


def get_options() -> Options:
    """返回 Pythonista 专用配置。"""
    return Options(
        excel=get_script_dir() / PYTHONISTA_EXCEL_FILENAME,
        start_row=PYTHONISTA_START_ROW,
        separator=PYTHONISTA_SEPARATOR,
        cookie=PYTHONISTA_COOKIE.strip(),
        timeout=PYTHONISTA_TIMEOUT,
        sleep=PYTHONISTA_SLEEP,
        no_backup=not PYTHONISTA_CREATE_BACKUP,
    )


def api_get(session: Any, url: str, params: dict[str, Any], timeout: float) -> dict[str, Any]:
    response = session.get(url, params=params, timeout=timeout)
    response.raise_for_status()
    payload = response.json()
    if payload.get("code") != 0:
        raise BilibiliAPIError(f"{url} 返回 code={payload.get('code')}, message={payload.get('message')}")
    return payload


def get_first_cid(session: Any, bvid: str, timeout: float) -> int:
    payload = api_get(session, VIEW_URL, {"bvid": bvid}, timeout)
    pages = payload.get("data", {}).get("pages") or []
    if not pages or not pages[0].get("cid"):
        raise BilibiliAPIError(f"未查询到 cid: {bvid}")
    return int(pages[0]["cid"])


def get_tags(session: Any, bvid: str, cid: int, timeout: float) -> list[dict[str, Any]]:
    payload = api_get(session, TAG_URL, {"bvid": bvid, "cid": cid}, timeout)
    data = payload.get("data")
    if not isinstance(data, list):
        raise BilibiliAPIError(f"TAG 接口返回格式异常: {bvid}: data 不是列表")
    return data


def extract_bgm_tag_name(tag_name: str) -> str:
    """只保留 BGM TAG 名称中《》内的内容。"""
    match = BOOK_TITLE_PATTERN.search(tag_name)
    if not match:
        return ""
    return match.group(1).strip()


def get_bgm_tags(tags: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """筛选 tag_type 为 bgm 的 TAG。"""
    return [tag for tag in tags if tag.get("tag_type") == "bgm"]


def make_session(cookie: str) -> Any:
    requests = __import__("requests")
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
            "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
            "Referer": "https://www.bilibili.com/",
        }
    )
    if cookie:
        session.headers["Cookie"] = cookie
    return session


def run(options: Options) -> int:
    if not options.excel.exists():
        print(f"未找到 Excel 文件: {options.excel}")
        return 2

    openpyxl = __import__("openpyxl")
    workbook = openpyxl.load_workbook(options.excel)
    worksheet = workbook.active
    session = make_session(options.cookie)

    processed = 0
    failed: list[tuple[int, str, str]] = []

    for row in range(options.start_row, worksheet.max_row + 1):
        raw_bvid = worksheet[f"{BVID_COLUMN}{row}"].value
        bvid = str(raw_bvid).strip() if raw_bvid is not None else ""
        if not bvid:
            continue

        try:
            cid = get_first_cid(session, bvid, options.timeout)
            tags = get_tags(session, bvid, cid, options.timeout)
            bgm_tags = get_bgm_tags(tags)
            tag_names = [
                name
                for tag in bgm_tags
                if tag.get("tag_name")
                for name in [extract_bgm_tag_name(str(tag["tag_name"]))]
                if name
            ]
            music_ids = [str(tag.get("music_id", "")).strip() for tag in bgm_tags if tag.get("music_id")]
            worksheet[f"{TAG_NAME_COLUMN}{row}"] = options.separator.join(tag_names)
            worksheet[f"{MUSIC_ID_COLUMN}{row}"] = options.separator.join(music_ids)
            processed += 1
            print(f"row {row}: {bvid} -> {len(tag_names)} 个 BGM TAG, {len(music_ids)} 个 music_id")
        except Exception as exc:  # noqa: BLE001 - 单行失败后继续批量处理。
            failed.append((row, bvid, str(exc)))
            print(f"row {row}: {bvid} 失败: {exc}")

        if options.sleep > 0:
            time.sleep(options.sleep)

    if not options.no_backup:
        backup_path = options.excel.with_suffix(options.excel.suffix + ".bak.xlsx")
        shutil.copy2(options.excel, backup_path)
        print(f"已写入备份: {backup_path}")

    workbook.save(options.excel)
    print(f"已更新 {processed} 行,文件: {options.excel}")

    if failed:
        print("失败明细:")
        for row, bvid, reason in failed:
            print(f"  row {row} ({bvid}): {reason}")
        return 1

    return 0


def main() -> int:
    return run(get_options())


if __name__ == "__main__":
    raise SystemExit(main())
