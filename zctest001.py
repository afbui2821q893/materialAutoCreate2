#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zctest001.py —— 料号申请自动化采集（Playwright 同步版）

相对原异步版的变化：
    1. 浏览器自动化改用 Playwright 同步 API（sync_playwright），
       所有 async def / await 去掉；流程本身严格串行，性能完全一致。
    2. 搜索按钮定位修复：不能用 get_by_role(name=re.compile(r"搜\\s*索"))，
       该正则是“包含匹配”，会同时命中两个「高级搜索」按钮，触发
       strict mode violation。这里锚定整名并限定主按钮 ant-btn-primary。
    3. 详情页解析容错：字段整列匹配不到（如分类/来源码未填）时按空值补齐，
       不再因列数为 0 被误判为“字段列数不一致”而丢弃整单。
    4. 详情页打开方式兼容：优先按新标签页处理，若 OA 在当前标签页内跳转，
       解析完后自动返回列表继续下一条。
    5. Excel 写回改为直接调用（原 asyncio.to_thread 只是把阻塞调用挪出
       事件循环，同步版没有事件循环要保护，直接写即可）；COM 在主线程
       初始化/释放。
    6. 搜索后等待列表内容真正刷新再读取，修复把搜索前的待办列表
       当成搜索结果的问题。
    7. 其余采集、清洗、统计逻辑保持不变。

依赖安装：
    pip install playwright pywin32
    playwright install chromium

用法：
    set OA_USERNAME=xxx& set OA_PASSWORD=yyy& python zctest001.py
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

import pythoncom
import win32com.client as win32
from playwright.sync_api import TimeoutError as PWTimeoutError
from playwright.sync_api import sync_playwright

# ------------------------------------------------------------------ 配置 --

BASE_DIR = Path(r"D:\pythonProjects\materialAutoCreate2")
TEMPLATE_XLS = BASE_DIR / "template.xls"       # 含 CGJ 工作表的模板
OUTPUT_XLS = BASE_DIR / "toBeCreated.xls"      # ERP 导入用输出文件（.xls）
SHEET_NAME = "CGJ"

OA_URL = (
    "http://191.168.5.50:8001/wui/index.html#/main/workflow/listDoing"
    "?menuIds=1,13&menuPathIds=1,13"           # 原 URL 里的 _key 是 UI 状态参数，已去掉
)
PROCESS_TITLE = "料号申请"
MACHINE_MODEL = "CZ采购件"

HEADLESS = os.environ.get("HEADLESS", "0") == "1"   # 默认有头，便于观察排查

PAGE_TIMEOUT_MS = 30_000        # 常规操作超时
POPUP_TIMEOUT_MS = 5_000        # 登录提示浮层等可选项
DETAIL_TIMEOUT_MS = 20_000      # 详情页字段渲染
NEW_TAB_TIMEOUT_MS = 15_000     # 点击行后等待详情出现

# 待办列表中每行第 2 列（请求标题列），读取列表统一用它
ROW_CELL_SELECTOR = "tbody.ant-table-tbody tr.ant-table-row td:nth-child(2)"

# 详情页明细表字段 -> 候选 CSS 选择器列表（按顺序尝试，取第一个非空结果）。
# 分类列原来只写了 `a`，当分类为空或不是链接渲染时匹配数为 0，会被列数校验
# 误判为“字段列数不一致”而整单丢弃；这里补上单元格文本兜底。
DETAIL_FIELDS = {
    "name":           ("td.detail_1_3_3.etype_3 > div > span > span",),
    "classification": ("td.detail_1_3_4.etype_3 a",
                       "td.detail_1_3_4.etype_3 > div > span > span",
                       "td.detail_1_3_4.etype_3"),
    "specification":  ("td.detail_1_3_5.etype_3 > div > span > span",),
    "brand":          ("td.detail_1_3_6.etype_3 > div > span > span",),
    "unit":           ("td.detail_1_3_7.etype_3 > div > span > span",),
    "source_code":    ("td.detail_1_3_8.etype_3 > div > span > span.child-item.wdb",),
    "weight":         ("td.detail_1_3_9.etype_3 > div > span > span",),
    "material":       ("td.detail_1_3_10.etype_3 > div > span > span",),
    "remark":         ("td.detail_1_3_11.etype_3 > div > span > span",),
    "instruction":    ("td.detail_1_3_12.etype_3 > div > span > span",),
}
# 有附件则整单跳过（对应原 field11036 附件字段）
ATTACHMENT_SELECTOR = "div.field11036_swapDiv .wea-upload-list-item-filesize"

# 单位映射表：页面单位值（为其子串）-> 表尾 [CODE] 中的标准单位码
UNIT_LIST = [
    "件（本、批、盆、束、项、次、个、件、只、支、栋、辆）[PCS]",
    "台[SET]",
    "套[SUIT]",
    "升[L]",
    "条（根）[SLIP]",
    "张(块、板)[SHEET]",
    "箱[BOX]",
    "卷[ROL]",
    "管[TUBE]",
    "瓶（罐）[BOT]",
    "盒[CAS]",
    "袋（包）[BAG]",
    "桶[BU]",
    "米[M]",
    "克[G]",
    "千克[KG]",
    "平方米[M2]",
    "毫米[MM]",
    "排[ROW]",
    "组[GROUP]",
    "小时[H]",
    "毫升[ML]",
    "吨[T]",
    "瓦[W]",
    "双[SHA]",
]

HEADER_CN = ["名称", "分类", "规格", "品牌", "单位", "来源码", "机型", "重量", "材质", "备注"]
HEADER_EN = ["name$", "classification$", "Specification", "BRAND", "UOM",
             "SourceType", "MRPJX", "Weight", "Material", "Remark"]

log = logging.getLogger("zctest001")


# ------------------------------------------------------------------ 数据 --

@dataclass
class MaterialRow:
    """一行物料（与输出 Excel 的 10 列一一对应）。"""
    name: str
    classification: str
    specification: str
    brand: str
    unit: str
    source_code: str
    machine_model: str
    weight: str
    material: str
    remark: str

    def to_list(self) -> list:
        return [self.name, self.classification, self.specification, self.brand,
                self.unit, self.source_code, self.machine_model, self.weight,
                self.material, self.remark]


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


def clean_text(text: str, space_replacement: str) -> str:
    """先全角转半角，再做符号替换（修复原脚本先替换后转换、
    导致全角空格和全角＊漏处理的顺序问题）。"""
    text = q2b_str(text.strip())
    return text.replace(" ", space_replacement).replace("*", "x")


def normalize_unit(unit: str) -> str:
    """页面单位 -> 标准单位码；无法识别时保留原值并告警（原来是静默写 None）。"""
    unit = q2b_str(unit.strip())
    if unit in ("g", "G", "克"):
        return "G"
    for pattern in UNIT_LIST:
        if unit in pattern:
            return re.search(r"[A-Za-z]+", pattern).group()
    log.warning("  单位无法识别: %r，按原值写入", unit)
    return unit


def normalize_source_code(source: str) -> str:
    """截取来源码中的英文字母段。

    原脚本 re.findall(r'[a-zA-Z]', ...)[0] 只取第一个字母，多字母来源码会被
    截断；这里改为取完整字母段。若业务确实只要首字母，改回 m.group()[0] 即可。
    """
    source = q2b_str(source.strip())
    m = re.search(r"[A-Za-z]+", source)
    if not m:
        log.warning("  来源码中未找到字母: %r，按原值写入", source)
        return source
    return m.group()


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

    # 兼容「搜 索 / 搜索」两种写法。
    # 注意：Playwright 的 name 正则按“包含”匹配，r"搜\s*索" 会同时命中
    # 「高级搜索」，触发 strict mode violation（resolved to 3 elements）。
    # 这里用锚定整名匹配排除「高级搜索」，并优先限定为主按钮 ant-btn-primary。
    search_btn = page.locator("button.ant-btn-primary").filter(
        has_text=re.compile(r"^\s*搜\s*索\s*$"))
    if search_btn.count() == 0:
        search_btn = page.get_by_role(
            "button", name=re.compile(r"^\s*搜\s*索\s*$"))
    # 点击搜索前先记录当前列表内容：搜索结果出来之前表格里仍是旧数据，
    # 原来的"等表格存在/网络空闲"会立即满足，导致把搜索前的待办列表当成结果
    # （曾出现实际 2 条待处理、却读到 10 条的情况）。
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
    detail_name_selector = DETAIL_FIELDS["name"][0].split(">")[0].strip()
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
        if page.locator(detail_name_selector).count() > 0:
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


def parse_detail(detail) -> list[MaterialRow] | None:
    """解析一条申请的明细表。

    返回 None 表示带附件（按原逻辑整单跳过）；
    名称列有数据而其他字段列数对不上时抛 ValueError，由调用方按失败处理。
    """
    # 等明细表渲染完成（名称列出现），替代原来的固定 sleep
    detail.wait_for_selector(DETAIL_FIELDS["name"][0].split(">")[0].strip(),
                             timeout=DETAIL_TIMEOUT_MS)

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

    values: dict[str, list[str]] = {
        "name": pick(DETAIL_FIELDS["name"], 0)}     # 名称列先确定数据行数
    expected = len(values["name"])
    if expected == 0:
        log.info("  明细表无数据行，跳过")
        return []

    for key, selectors in DETAIL_FIELDS.items():
        if key == "name":
            continue
        texts = pick(selectors, expected)
        if not texts:                # 整列为空（如分类未填）→ 按空串补齐，不整单丢弃
            log.warning("  字段 %s 未匹配到内容，按空值写入", key)
            texts = [""] * expected
        elif len(texts) != expected:
            lengths = {k: len(v) for k, v in values.items()}
            lengths[key] = len(texts)
            raise ValueError(f"字段列数不一致: {lengths}")
        values[key] = texts

    rows: list[MaterialRow] = []
    for i in range(expected):
        remark = values["remark"][i].strip()
        instruction = values["instruction"][i].strip()
        if remark == "" and instruction != "":
            remark = instruction

        rows.append(MaterialRow(
            name=clean_text(values["name"][i], ","),
            classification=q2b_str(values["classification"][i].strip()),
            specification=clean_text(values["specification"][i], "_"),
            brand=clean_text(values["brand"][i], "_"),
            unit=normalize_unit(values["unit"][i]),
            source_code=normalize_source_code(values["source_code"][i]),
            machine_model=MACHINE_MODEL,
            weight=q2b_str(values["weight"][i].strip()),
            material=q2b_str(values["material"][i].strip()),
            remark=q2b_str(remark),
        ))
    return rows


# ------------------------------------------------------------------ Excel --

def write_excel(rows: list[MaterialRow]) -> None:
    """同步写 Excel（COM 调用，主线程直接执行，不再绕线程）。"""
    if not TEMPLATE_XLS.exists():
        raise FileNotFoundError(f"模板不存在: {TEMPLATE_XLS}")

    pythoncom.CoInitialize()      # 主线程同样显式初始化，避免 COM 未初始化
    excel = None
    wb = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(TEMPLATE_XLS))
        sheet = wb.Worksheets(SHEET_NAME)

        # 两行表头写到 A1:J2（原脚本 A1:J3 与 2 行表头不匹配），数据从 A3 起
        sheet.Range("A1:J2").Value = (HEADER_CN, HEADER_EN)
        data = [row.to_list() for row in rows]
        if data:
            sheet.Range(f"A3:J{len(data) + 2}").Value = data

        if OUTPUT_XLS.exists():
            OUTPUT_XLS.unlink()   # 不存在时不会像原来那样直接崩
        wb.SaveAs(str(OUTPUT_XLS), FileFormat=56)      # 56 = .xls
        log.info("已生成: %s（%d 行数据）", OUTPUT_XLS, len(data))
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()          # 兜底退出，避免 EXCEL.EXE 残留锁文件
        pythoncom.CoUninitialize()


# ------------------------------------------------------------------ 主流程 --

def run(username: str, password: str) -> int:
    rows: list[MaterialRow] = []
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
                        log.info("  解析到 %d 行物料", len(parsed))
                except PWTimeoutError:
                    stats["failed"] += 1
                    log.error("  详情页加载超时，跳过")
                except Exception:
                    stats["failed"] += 1
                    log.exception("  解析详情页失败，跳过")
                finally:
                    if is_new_tab:
                        detail.close()      # 修复原脚本的标签页泄漏
                    else:
                        back_to_list(page)  # 当前页跳转：返回列表继续下一条
        finally:
            browser.close()

    log.info("统计: 申请 %d 条，带附件跳过 %d，失败 %d，共采集物料 %d 行",
             stats["requests"], stats["skipped"], stats["failed"], len(rows))

    if not rows:
        log.warning("未采集到任何物料数据，仅生成表头文件")
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
