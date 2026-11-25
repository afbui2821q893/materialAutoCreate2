#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件的环境变量配置：
PYTHONUNBUFFERED=1;OA_USERNAME=CJ6149;OA_PASSWORD=oaOA1354648!2
zctest003.py —— 料号修改申请自动化采集（Playwright 同步版）

仿 zctest001.py（料号申请）的结构，采集「料号修改申请」待处理流程的明细。
与 001 的差异：
    1. 流程标题改为「料号修改申请」。
    2. 详情页字段分两类：
         - 顶部单值字段：申请人 / 申请日期 / 类型（每条申请一份，逐行复制）
         - 底部明细表字段：序号 / 物料编码 / 修改前主体项 / 修改后主体项 /
                           修改原因 / 备注
       合计 9 列，与输出 Excel 一一对应。
    3. 输出不再基于 template.xls，直接用 openpyxl 新建工作簿写入 xlsx。
    4. 搜索后等待列表真正刷新、异步字段短时重读等容错逻辑与 001 保持一致。
    5. 带附件的申请整单跳过（附件字段为 field11823）。

依赖安装：
    pip install playwright openpyxl
    playwright install chromium

用法：
    set OA_USERNAME=xxx& set OA_PASSWORD=yyy& python zctest003.py
    set HEADLESS=1  可切换无头模式
"""

from __future__ import annotations

import getpass
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import TimeoutError as PWTimeoutError
from playwright.sync_api import sync_playwright

# ------------------------------------------------------------------ 配置 --

BASE_DIR = Path(r"D:\pythonProjects\materialAutoCreate2")
OUTPUT_XLSX = BASE_DIR / "toBeModified.xlsx"   # 输出文件（自行创建，不依赖模板）

OA_URL = (
    "http://191.168.5.50:8001/wui/index.html#/main/workflow/listDoing"
    "?menuIds=1,13&menuPathIds=1,13"           # 原 URL 里的 _key 是 UI 状态参数，已去掉
)
PROCESS_TITLE = "料号修改申请"

HEADLESS = os.environ.get("HEADLESS", "0") == "1"   # 默认有头，便于观察排查

PAGE_TIMEOUT_MS = 30_000        # 常规操作超时
POPUP_TIMEOUT_MS = 5_000        # 登录提示浮层等可选项
DETAIL_TIMEOUT_MS = 20_000      # 详情页字段渲染
NEW_TAB_TIMEOUT_MS = 15_000     # 点击行后等待详情出现

# 列表页每行的第 2 列（流程标题），读取列表统一用它
ROW_CELL_SELECTOR = "tbody.ant-table-tbody tr.ant-table-row td:nth-child(2)"

# 顶部单值字段（每条申请只有一个值，泛微用 mainTd_X_Y 命名）
HEADER_FIELDS = {
    "applicant":   ("td.mainTd_5_4.etype_3 a.child-item.wdb",
                     "td.mainTd_5_4.etype_3 span.child-item.wdb",
                     "td.mainTd_5_4.etype_3"),
    "apply_date":  ("td.mainTd_5_6.etype_3 span.child-item.wdb",
                     "td.mainTd_5_6.etype_3 > div > span > span",
                     "td.mainTd_5_6.etype_3"),
    "modify_type": ("td.mainTd_6_2.etype_3 span.child-item.wdb",
                     "td.mainTd_6_2.etype_3 > div > span > span",
                     "td.mainTd_6_2.etype_3"),
}

# 底部明细表字段（逐行多值，泛微用 detail_1_3_X 命名）。
# 物料编码作为确定数据行数的锚点（对应 001 的 name 列）。
DETAIL_FIELDS = {
    "material_code": ("td.detail_1_3_2.etype_3 > div > span > span",
                       "td.detail_1_3_2.etype_3"),
    "seq":           ("td.detail_1_3_1.etype_22 > div > span",
                       "td.detail_1_3_1.etype_22"),
    "before_item":   ("td.detail_1_3_3.etype_3 > div > span > span",
                       "td.detail_1_3_3.etype_3"),
    "after_item":    ("td.detail_1_3_4.etype_3 > div > span > span",
                       "td.detail_1_3_4.etype_3"),
    "reason":        ("td.detail_1_3_5.etype_3 > div > span > span",
                       "td.detail_1_3_5.etype_3"),
    "remark":        ("td.detail_1_3_6.etype_3 > div > span > span",
                       "td.detail_1_3_6.etype_3"),
}

# 等明细表渲染：物料编码列出现即认为明细行已开始渲染
DETAIL_ANCHOR_SELECTOR = "td.detail_1_3_2.etype_3"
# 带附件则整单跳过（对应 field11823 附件字段 xianggfj）
ATTACHMENT_SELECTOR = "div.field11823_swapDiv .wea-upload-list-item-filesize"

# 必须读到值才继续的字段（备注允许为空）。申请人/申请日期/类型以及明细的
# 主体项等由前端异步渲染，锚点列出现时它们可能还没填值，读到整列为空时短时重读。
REQUIRED_HEADER_FIELDS = ("applicant", "apply_date", "modify_type")
REQUIRED_DETAIL_FIELDS = ("seq", "material_code", "before_item", "after_item", "reason")
FIELD_RETRY_TIMEOUT_MS = 6_000   # 异步字段重读的最长等待

OUTPUT_HEADERS = ["申请人", "申请日期", "类型", "序号", "物料编码",
                  "修改前主体项", "修改后主体项", "修改原因", "备注"]

log = logging.getLogger("zctest003")


# ------------------------------------------------------------------ 数据 --

@dataclass
class ModifyRow:
    """一行修改明细（与输出 Excel 的 9 列一一对应）。"""
    applicant: str
    apply_date: str
    modify_type: str
    seq: str
    material_code: str
    before_item: str
    after_item: str
    reason: str
    remark: str

    def to_list(self) -> list:
        return [self.applicant, self.apply_date, self.modify_type, self.seq,
                self.material_code, self.before_item, self.after_item,
                self.reason, self.remark]


# -------------------------------------------------------------- 清洗工具 --

def q2b(ch: str) -> str:
    """单字符全角转半角；转换结果不在可见 ASCII 范围则原样返回。"""
    code = ord(ch)
    if code == 0x3000:            # 全角空格
        return " "
    code -= 0xFEE0
    if 0x20 <= code <= 0x7E:
        return chr(code)
    return ch


def q2b_str(text: str) -> str:
    return "".join(q2b(ch) for ch in text)


def clean_field(text: str) -> str:
    """统一清洗：全角转半角后去首尾空白（含物料编码前导的 &nbsp;）。"""
    return q2b_str(text).strip()


# -------------------------------------------------------------- 页面交互 --

def login(page, username: str, password: str) -> None:
    log.info("打开 OA 并登录")
    page.goto(OA_URL, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
    page.fill("#loginid", username)
    page.fill("#userpassword", password)
    page.click("#submit")
    page.wait_for_load_state("domcontentloaded")

    # 关闭登录后的提示浮层（不一定出现，超时即忽略）
    try:
        page.locator(".anticon-cross").first.click(timeout=POPUP_TIMEOUT_MS)
    except PWTimeoutError:
        log.debug("未出现提示浮层，跳过关闭")
    page.reload(wait_until="domcontentloaded")


def search_pending_requests(page) -> None:
    log.info("高级搜索: %s / 待处理", PROCESS_TITLE)
    # 页面可能同时渲染多个「高级搜索」入口，显式取第一个
    page.locator(".wea-advanced-search").first.click()
    page.fill("#requestname--1", PROCESS_TITLE)

    # 展开状态下拉并选择「待处理」
    page.click("#weaSelect_3 .ant-select-selection-selected-value")
    page.locator("li[title='待处理']").first.click()

    # 锚定整名匹配排除「高级搜索」，并优先限定为主按钮 ant-btn-primary
    search_btn = page.locator("button.ant-btn-primary").filter(
        has_text=re.compile(r"^\s*搜\s*索\s*$"))
    if search_btn.count() == 0:
        search_btn = page.get_by_role(
            "button", name=re.compile(r"^\s*搜\s*索\s*$"))
    # 点击搜索前先记录当前列表内容：搜索结果出来之前表格里仍是旧数据，
    # 直接等表格存在/网络空闲会立即满足，导致把搜索前的待办列表当成结果。
    before = [t.strip() for t in page.locator(ROW_CELL_SELECTOR).all_text_contents()]

    search_btn.first.click()

    deadline = time.monotonic() + DETAIL_TIMEOUT_MS / 1000
    while True:
        # 等表格 loading 结束（没有 loading 时立即返回）
        try:
            page.wait_for_selector(".ant-spin-spinning", state="hidden", timeout=1000)
        except PWTimeoutError:
            pass
        current = [t.strip() for t in page.locator(ROW_CELL_SELECTOR).all_text_contents()]
        empty = page.locator(".ant-table-placeholder").count() > 0
        # 内容确实变化（或确认为空表）才算刷新完成；
        # 读到空列表但无占位符说明还在渲染中，继续等。
        if current != before and (current or empty):
            log.info("搜索结果已刷新，共 %d 条", len(current))
            return
        if time.monotonic() >= deadline:
            log.warning("等待搜索结果刷新超时，按当前列表继续（结果可能与搜索前一致）")
            return
        page.wait_for_timeout(300)


def open_detail(context, page, cell) -> tuple[object, bool]:
    """点击列表行打开详情页。

    返回 (detail, is_new_tab)。点击目标优先取单元格里的标题链接：
    td 往往比文字宽得多，点 td 中心容易落在空白处，导致“点了没反应”。
    点击后轮询 新标签页 / URL 变化 / 明细表出现 三种结果；
    5 秒仍无反应时补一次双击（部分 OA 列表是双击打开）。
    """
    link = cell.locator("a").first
    target = link if link.count() > 0 else cell
    log.info("  点击行: %s（%s）",
             cell.inner_text().strip().replace("\n", " ")[:40],
             "标题链接" if target is link else "单元格")

    pages_before = len(context.pages)
    url_before = page.url
    start = time.monotonic()
    dbl_clicked = False
    target.click()

    while time.monotonic() - start < NEW_TAB_TIMEOUT_MS / 1000:
        if len(context.pages) > pages_before:
            log.info("  详情页已在新标签页打开")
            return context.pages[-1], True
        if page.url != url_before:
            log.info("  详情页在当前标签页打开")
            return page, False
        if page.locator(DETAIL_ANCHOR_SELECTOR).count() > 0:
            log.info("  详情页在当前标签页打开")
            return page, False

        if not dbl_clicked and time.monotonic() - start >= 5:
            log.warning("  单击无响应，尝试双击")
            target.dblclick()
            dbl_clicked = True
        page.wait_for_timeout(500)

    raise PWTimeoutError(
        f"点击后无响应（pages={len(context.pages)}，url={page.url}）")


def back_to_list(page) -> None:
    """详情在当前标签页打开后返回列表；表格没恢复时重新搜索。"""
    try:
        page.go_back(wait_until="domcontentloaded")
        page.wait_for_selector(
            "tbody.ant-table-tbody tr.ant-table-row, .ant-table-placeholder",
            timeout=DETAIL_TIMEOUT_MS,
        )
    except PWTimeoutError:
        log.warning("  返回列表失败，重新执行搜索")
        try:
            search_pending_requests(page)
        except Exception:
            log.exception("  重新搜索失败，后续申请可能无法处理")


def parse_detail(detail) -> list[ModifyRow] | None:
    """解析一条「料号修改申请」的详情页。

    返回 None 表示带附件（整单跳过）；返回 [] 表示无明细行。
    顶部字段（申请人/申请日期/类型）每条申请一份，逐行复制到明细行上。
    """
    # 等明细表渲染完成（物料编码列出现），替代原来的固定 sleep
    detail.wait_for_selector(DETAIL_ANCHOR_SELECTOR, timeout=DETAIL_TIMEOUT_MS)

    if detail.locator(ATTACHMENT_SELECTOR).count() > 0:
        log.info("  该申请带附件，跳过")
        return None

    def pick(selectors: tuple[str, ...], expect: int) -> list[str]:
        """按候选顺序取列文本：优先列数与 expect 一致的选择器，否则取第一个非空。"""
        fallback: list[str] | None = None
        for selector in selectors:
            texts = detail.locator(selector).all_text_contents()
            if not texts:
                continue
            if expect <= 0 or len(texts) == expect:
                return texts
            fallback = fallback or texts
        return fallback if fallback is not None else []

    def read_header() -> tuple[dict[str, list[str]], list[str]]:
        values: dict[str, list[str]] = {}
        not_ready: list[str] = []
        for key, selectors in HEADER_FIELDS.items():
            texts = pick(selectors, 1)
            values[key] = texts
            if key in REQUIRED_HEADER_FIELDS and (
                    len(texts) != 1 or not texts[0].strip()):
                not_ready.append(key)
        return values, not_ready

    def read_detail() -> tuple[dict[str, list[str]], list[str], int]:
        values: dict[str, list[str]] = {
            "material_code": pick(DETAIL_FIELDS["material_code"], 0)}  # 先确定行数
        expected_ = len(values["material_code"])
        not_ready: list[str] = []
        for key, selectors in DETAIL_FIELDS.items():
            if key == "material_code":
                continue
            texts = pick(selectors, expected_)
            values[key] = texts
            if key in REQUIRED_DETAIL_FIELDS and (
                    len(texts) != expected_ or all(not t.strip() for t in texts)):
                not_ready.append(key)
        return values, not_ready, expected_

    # 申请人/申请日期/类型以及明细主体项等由前端异步渲染：锚点列出现时它们的值
    # 可能还没填上，直接读会整列为空。这里对必填字段短时重读。
    deadline = time.monotonic() + FIELD_RETRY_TIMEOUT_MS / 1000
    header_values: dict[str, list[str]] = {}
    detail_values: dict[str, list[str]] = {}
    expected = 0
    h_not_ready: list[str] = []
    d_not_ready: list[str] = []
    while True:
        header_values, h_not_ready = read_header()
        detail_values, d_not_ready, expected = read_detail()
        if expected == 0:
            log.info("  明细表无数据行，跳过")
            return []
        if not (h_not_ready or d_not_ready) or time.monotonic() >= deadline:
            break
        detail.wait_for_timeout(300)
    if h_not_ready or d_not_ready:
        log.warning("  字段 %s 等待后仍为空，按空值写入",
                    ", ".join(h_not_ready + d_not_ready))

    # 明细列补齐/校验
    for key in DETAIL_FIELDS:
        texts = detail_values.get(key, [])
        if not texts:                # 整列为空（如备注未填）→ 按空串补齐，不整单丢弃
            log.warning("  字段 %s 未匹配到内容，按空值写入", key)
            texts = [""] * expected
        elif len(texts) != expected:
            lengths = {k: len(v) for k, v in detail_values.items()}
            lengths[key] = len(texts)
            raise ValueError(f"字段列数不一致: {lengths}")
        detail_values[key] = texts

    # 顶部字段取第一个值（每条申请一份）
    header = {k: clean_field(v[0] if v else "") for k, v in header_values.items()}

    rows: list[ModifyRow] = []
    for i in range(expected):
        rows.append(ModifyRow(
            applicant=header["applicant"],
            apply_date=header["apply_date"],
            modify_type=header["modify_type"],
            seq=clean_field(detail_values["seq"][i]),
            material_code=clean_field(detail_values["material_code"][i]),
            before_item=clean_field(detail_values["before_item"][i]),
            after_item=clean_field(detail_values["after_item"][i]),
            reason=clean_field(detail_values["reason"][i]),
            remark=clean_field(detail_values["remark"][i]),
        ))
    return rows


# ------------------------------------------------------------------ Excel --

def write_excel(rows: list[ModifyRow]) -> None:
    """直接用 openpyxl 新建工作簿写入 xlsx（不依赖模板）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "料号修改"

    # 表头
    ws.append(OUTPUT_HEADERS)
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 数据行
    for row in rows:
        ws.append(row.to_list())

    # 列宽自适应（按内容长度估算），并冻结首行
    for col_idx, _ in enumerate(OUTPUT_HEADERS, start=1):
        max_len = len(OUTPUT_HEADERS[col_idx - 1])
        for row in rows:
            val = str(row.to_list()[col_idx - 1] or "")
            max_len = max(max_len, len(val))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(
            max(max_len + 2, 10), 40)
    ws.freeze_panes = "A2"

    if OUTPUT_XLSX.exists():
        OUTPUT_XLSX.unlink()
    wb.save(str(OUTPUT_XLSX))
    log.info("已生成: %s（%d 行数据）", OUTPUT_XLSX, len(rows))


# ------------------------------------------------------------------ 主流程 --

def run(username: str, password: str) -> int:
    rows: list[ModifyRow] = []
    stats = {"requests": 0, "skipped": 0, "failed": 0}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=HEADLESS)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(PAGE_TIMEOUT_MS)

        try:
            login(page, username, password)
            search_pending_requests(page)

            cells = page.locator(ROW_CELL_SELECTOR)
            count = cells.count()
            log.info("待处理申请共 %d 条", count)

            for idx in range(count):
                stats["requests"] += 1
                log.info("处理第 %d/%d 条申请", idx + 1, count)

                # 点击行打开详情：优先新标签页，兼容当前页内跳转
                try:
                    detail, is_new_tab = open_detail(context, page, cells.nth(idx))
                except PWTimeoutError:
                    stats["failed"] += 1
                    log.error("  点击后未打开详情页，跳过")
                    continue

                try:
                    detail.wait_for_load_state("domcontentloaded")
                    parsed = parse_detail(detail)
                    if parsed is None:
                        stats["skipped"] += 1
                    else:
                        rows.extend(parsed)
                        log.info("  解析到 %d 行修改明细", len(parsed))
                except PWTimeoutError:
                    stats["failed"] += 1
                    log.error("  详情页加载超时，跳过")
                except Exception:
                    stats["failed"] += 1
                    log.exception("  解析详情页失败，跳过")
                finally:
                    if is_new_tab:
                        detail.close()      # 修复标签页泄漏
                    else:
                        back_to_list(page)  # 当前页跳转：返回列表继续下一条
        finally:
            browser.close()

    log.info("统计: 申请 %d 条，带附件跳过 %d，失败 %d，共采集修改明细 %d 行",
             stats["requests"], stats["skipped"], stats["failed"], len(rows))

    if not rows:
        log.warning("未采集到任何修改明细，仅生成表头文件")
    write_excel(rows)
    return 0


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    username = os.environ.get("OA_USERNAME", "").strip() or input("OA 账号: ").strip()
    password = os.environ.get("OA_PASSWORD", "").strip() or getpass.getpass("OA 密码: ")
    if not username or not password:
        log.error("未提供账号或密码（可设置环境变量 OA_USERNAME / OA_PASSWORD）")
        return 1
    try:
        return run(username, password)
    except Exception:
        log.exception("脚本执行失败")
        return 1


if __name__ == "__main__":
    sys.exit(main())
