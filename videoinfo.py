#!/usr/bin/env python3
"""批量将 B 站视频 TAG 名称和 music_id 写入 Excel 文件。

脚本会从 A 列读取 BVID,先查询每个视频的 CID,再调用支持 BGM 信息的
B 站 TAG 接口。TAG 名称会写入原工作簿的 R 列,music_id 会写入 S 列。

桌面命令行用法:
    python videoinfo.py videos.xlsx
    python videoinfo.py videos.xlsx --cookie "SESSDATA=xxx"
    python videoinfo.py videos.xlsx --sheet Sheet1 --start-row 2

Pythonista 用法:
    1. 在 Pythonista 中运行本脚本。
    2. 按提示选择 .xlsx 文件并填写可选参数。
    3. 脚本会在保存前自动创建 .bak.xlsx 备份,除非关闭备份选项。

依赖:
    pip install openpyxl requests

Pythonista 需要先安装 requests 和 openpyxl。可在 StaSh 中执行:
    pip install requests openpyxl
"""

from __future__ import annotations

import argparse
import importlib.util
import shutil
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any


VIEW_URL = "https://api.bilibili.com/x/web-interface/view"
TAG_URL = "https://api.bilibili.com/x/web-interface/view/detail/tag"

BVID_COLUMN = "A"
TAG_NAME_COLUMN = "R"
MUSIC_ID_COLUMN = "S"


def is_pythonista() -> bool:
    """判断当前脚本是否运行在 Pythonista 环境。"""
    return sys.platform == "ios" and importlib.util.find_spec("dialogs") is not None


@dataclass
class Options:
    """脚本运行参数。"""

    excel: Path
    sheet: str | None = None
    start_row: int = 2
    separator: str = "; "
    cookie: str = ""
    timeout: float = 10.0
    sleep: float = 0.3
    no_backup: bool = False


class BilibiliAPIError(RuntimeError):
    """B 站接口返回错误时抛出。"""


def parse_args() -> Options:
    parser = argparse.ArgumentParser(
        description="从 Excel 的 A 列读取 BVID,并把 tag_name/music_id 写入 R/S 列。"
    )
    parser.add_argument("excel", type=Path, help="需要原地更新的 .xlsx 文件路径。")
    parser.add_argument("--sheet", help="工作表名称。默认使用当前活动工作表。")
    parser.add_argument("--start-row", type=int, default=2, help="第一行数据所在行号。默认:2。")
    parser.add_argument(
        "--separator", default="; ", help="单个视频存在多个 TAG 时使用的分隔符。默认:'; '。"
    )
    parser.add_argument(
        "--cookie",
        default="",
        help="可选的 Cookie 请求头,例如 'SESSDATA=xxx'。参考接口可能需要登录态。",
    )
    parser.add_argument("--timeout", type=float, default=10.0, help="HTTP 超时时间,单位秒。默认:10。")
    parser.add_argument("--sleep", type=float, default=0.3, help="每个视频之间的请求间隔,单位秒。默认:0.3。")
    parser.add_argument(
        "--no-backup", action="store_true", help="覆盖工作簿前不创建 .bak.xlsx 备份。"
    )
    namespace = parser.parse_args()
    return Options(**vars(namespace))


def input_alert(title: str, message: str, default: str = "") -> str:
    """使用 Pythonista 弹窗读取文本。"""
    console = __import__("console")
    return console.input_alert(title, message, default, "确定", hide_cancel_button=False)


def ask_bool(title: str, message: str, default: bool) -> bool:
    """使用 Pythonista 弹窗读取布尔值。"""
    console = __import__("console")
    first = "是" if default else "否"
    second = "否" if default else "是"
    choice = console.alert(title, message, first, second, hide_cancel_button=True)
    return default if choice == 1 else not default


def parse_pythonista_options() -> Options:
    """在 Pythonista 中通过系统文件选择器和弹窗收集参数。"""
    dialogs = __import__("dialogs")
    picked_file = dialogs.pick_document(types=["public.data", "org.openxmlformats.spreadsheetml.sheet"])
    if not picked_file:
        raise SystemExit("已取消选择文件。")

    sheet = input_alert("工作表名称", "留空则使用当前活动工作表。", "").strip() or None
    start_row_text = input_alert("起始行", "第一行数据所在行号。", "2").strip() or "2"
    separator = input_alert("分隔符", "多个 TAG/music_id 的分隔符。", "; ")
    cookie = input_alert("Cookie", "可选,例如 SESSDATA=xxx；不需要可留空。", "")
    timeout_text = input_alert("超时秒数", "HTTP 请求超时时间。", "10").strip() or "10"
    sleep_text = input_alert("请求间隔", "每个视频之间等待的秒数。", "0.3").strip() or "0.3"
    create_backup = ask_bool("创建备份", "保存前是否创建 .bak.xlsx 备份？", True)

    return Options(
        excel=Path(picked_file),
        sheet=sheet,
        start_row=int(start_row_text),
        separator=separator,
        cookie=cookie.strip(),
        timeout=float(timeout_text),
        sleep=float(sleep_text),
        no_backup=not create_backup,
    )


def get_options() -> Options:
    """根据运行环境选择命令行参数或 Pythonista 交互式参数。"""
    if is_pythonista() and len(sys.argv) == 1:
        return parse_pythonista_options()
    return parse_args()


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


def make_session(cookie: str) -> Any:
    requests = __import__("requests")
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
            "Referer": "https://www.bilibili.com/",
        }
    )
    if cookie:
        session.headers["Cookie"] = cookie
    return session


def show_pythonista_result(title: str, message: str) -> None:
    """在 Pythonista 中弹窗提示最终结果。"""
    if is_pythonista():
        console = __import__("console")
        console.alert(title, message, "确定", hide_cancel_button=True)


def run(options: Options) -> int:
    if not options.excel.exists():
        print(f"未找到 Excel 文件: {options.excel}", file=sys.stderr)
        show_pythonista_result("运行失败", f"未找到 Excel 文件:\n{options.excel}")
        return 2

    openpyxl = __import__("openpyxl")
    workbook = openpyxl.load_workbook(options.excel)
    worksheet = workbook[options.sheet] if options.sheet else workbook.active
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
            tag_names = [str(tag.get("tag_name", "")).strip() for tag in tags if tag.get("tag_name")]
            music_ids = [str(tag.get("music_id", "")).strip() for tag in tags if tag.get("music_id")]
            worksheet[f"{TAG_NAME_COLUMN}{row}"] = options.separator.join(tag_names)
            worksheet[f"{MUSIC_ID_COLUMN}{row}"] = options.separator.join(music_ids)
            processed += 1
            print(f"row {row}: {bvid} -> {len(tag_names)} 个 TAG, {len(music_ids)} 个 music_id")
        except Exception as exc:  # noqa: BLE001 - 单行失败后继续批量处理。
            failed.append((row, bvid, str(exc)))
            print(f"row {row}: {bvid} 失败: {exc}", file=sys.stderr)

        if options.sleep > 0:
            time.sleep(options.sleep)

    if not options.no_backup:
        backup_path = options.excel.with_suffix(options.excel.suffix + ".bak.xlsx")
        shutil.copy2(options.excel, backup_path)
        print(f"已写入备份: {backup_path}")

    workbook.save(options.excel)
    print(f"已更新 {processed} 行,文件: {options.excel}")

    if failed:
        print("失败明细:", file=sys.stderr)
        for row, bvid, reason in failed:
            print(f"  row {row} ({bvid}): {reason}", file=sys.stderr)
        show_pythonista_result("部分失败", f"已更新 {processed} 行,{len(failed)} 行失败。详情请查看控制台。")
        return 1

    show_pythonista_result("运行完成", f"已更新 {processed} 行。\n文件:\n{options.excel}")
    return 0


def main() -> int:
    return run(get_options())


if __name__ == "__main__":
    raise SystemExit(main())
