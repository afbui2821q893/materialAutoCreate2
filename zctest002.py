#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zctest002.py —— 料号申请自动填号（Playwright 同步重构版，重构自 zytest002V1.py）

流程：
    1. 读取 1.txt，按 `数字+字母` 提取料号清单（如 1090112260A）
    2. 启动 Chromium 登录内网泛微 OA
    3. 高级搜索「料号申请」的待处理流程
    4. 逐条打开详情页：带附件的整单跳过，其余按明细行依次填入料号
    5. 详情页保持打开、不保存不提交，留给人工确认后手动提交
    6. 跑完后浏览器不自动关闭，脚本阻塞等待，用户手动关闭浏览器后退出

相对原脚本的变化：
    - Selenium + chromedriver 改为 Playwright，元素操作自带自动等待，
      彻底去掉 time.sleep（包括“物料超过 100 行要等 6 秒”的魔法数字）
    - 凭据不再硬编码：优先读环境变量 OA_USERNAME / OA_PASSWORD，
      缺失时交互式输入
    - 修复原脚本问题：料号不足时直接 IndexError、附件/输入框未渲染完就读取、
      `clear()+send_keys` 的输入方式、窗口句柄靠下标取导致切错页、
      行元素只抓一次容易失效、搜索按钮/高级搜索入口多元素命中
    - 填料号前先全选删除旧内容再写入，避免与页面已填值粘连
    - 明细行数先用网络空闲 + 输入框出现双重等待确认，再一次性填入，
      填入数量与料号清单做匹配校验，避免错位脏数据
    - 搜索后等待列表内容真正刷新再读取，修复把搜索前的待办列表
      当成搜索结果的问题
    - 全程 logging 输出统计，料号剩余/不足都会显式告警
    - 结束不再调用 browser.close()：保留会话与已填写页面，等待用户手动关闭

依赖安装：
    pip install playwright
    playwright install chromium

用法：
    set OA_USERNAME=xxx& set OA_PASSWORD=yyy& python zctest002.py
    set HEADLESS=1  可切换无头模式
"""

from __future__ import annotations

import getpass
import logging
import os
import re
import sys
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeoutError
from playwright.sync_api import sync_playwright

# ------------------------------------------------------------------ 配置 --

BASE_DIR = Path(r"D:\pythonProjects\materialAutoCreate2")
CODE_FILE = BASE_DIR / "1.txt"                 # 待填入的料号清单（每行一个或多个）

OA_URL = (
    "http://191.168.5.50:8001/wui/index.html#/main/workflow/listDoing"
    "?menuIds=1,13&menuPathIds=1,13"           # 原 URL 里的 _key 是 UI 状态参数，已去掉
)
PROCESS_TITLE = "料号申请"

HEADLESS = os.environ.get("HEADLESS", "0") == "1"   # 默认有头，便于观察

PAGE_TIMEOUT_MS = 30_000        # 常规操作超时
POPUP_TIMEOUT_MS = 5_000        # 登录提示浮层等可选项
DETAIL_TIMEOUT_MS = 20_000      # 详情页渲染 / 输入框出现
NEW_TAB_TIMEOUT_MS = 15_000     # 点击行后等待新标签页

# 详情页里料号输入框（原 XPath: //input[@class='wf-input wf-input-1
# wf-input-detail wf-input-field11039']，这里按 class token 匹配更稳）
CODE_INPUT_SELECTOR = "input.wf-input-field11039"
# 有附件则整单跳过（对应原 field11036 附件字段）
ATTACHMENT_SELECTOR = "div.field11036_swapDiv .wea-upload-list-item-filesize"
# 列表页每行的第 2 列（流程标题），点击可打开详情
ROW_CELL_SELECTOR = "tbody.ant-table-tbody tr.ant-table-row td:nth-child(2)"

log = logging.getLogger("zctest002")


# -------------------------------------------------------------- 料号清单 --

def load_material_codes(path: Path) -> list[str]:
    """从清单文件提取料号（保持原脚本的 `数字+字母` 规则与出现顺序）。"""
    if not path.exists():
        raise FileNotFoundError(f"料号清单不存在: {path}")

    codes: list[str] = []
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        m = re.search(r"\d+[A-Za-z]", line)
        if m:
            codes.append(m.group())
    return codes


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

    # Playwright 的 name 正则按“包含”匹配，r"搜\s*索" 会误命中「高级搜索」，
    # 这里锚定整名并限定主按钮；文案兼容「搜索 / 搜 索」
    search_btn = page.locator("button.ant-btn-primary").filter(
        has_text=re.compile(r"^\s*搜\s*索\s*$"))
    if search_btn.count() == 0:
        search_btn = page.get_by_role(
            "button", name=re.compile(r"^\s*搜\s*索\s*$"))
    # 点击搜索前先记录当前列表内容：搜索结果出来之前表格里仍是旧数据，
    # 原来的“等表格存在/网络空闲”会立即满足，导致把搜索前的待办列表当成结果
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

    返回 (detail, is_new_tab)。点击目标优先取标题链接：td 往往比文字宽，
    点 td 中心容易落在空白处，表现为“点击了但浏览器没反应”。
    点击后轮询 新标签页 / URL 变化 / 详情输入框出现；
    5 秒仍无反应时补一次双击（部分泛微列表是双击打开）。
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
        if (page.locator(CODE_INPUT_SELECTOR).count() > 0
                or page.locator(ATTACHMENT_SELECTOR).count() > 0):
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


def fill_detail(detail, codes: list[str], cursor: int) -> tuple[int, str]:
    """在一条申请的详情页填入料号。

    返回 (新游标, 状态)，状态为 filled / skipped（带附件）/ empty（无明细行）。
    料号不足时抛 ValueError，由调用方中止，避免错位写入。
    """
    # 替代原脚本固定 sleep(6)/sleep(2)：等页面数据加载完成
    try:
        detail.wait_for_load_state("networkidle", timeout=DETAIL_TIMEOUT_MS)
    except PWTimeoutError:
        log.warning("  networkidle 超时，继续尝试读取详情页")

    if detail.locator(ATTACHMENT_SELECTOR).count() > 0:
        log.info("  该申请带附件，跳过")
        return cursor, "skipped"

    try:
        detail.wait_for_selector(CODE_INPUT_SELECTOR, timeout=DETAIL_TIMEOUT_MS)
    except PWTimeoutError:
        log.warning("  未发现料号输入框（可能无明细行），跳过")
        return cursor, "empty"

    inputs = detail.locator(CODE_INPUT_SELECTOR)
    count = inputs.count()
    if cursor + count > len(codes):
        raise ValueError(
            f"料号清单不足：本单需要 {count} 个，剩余 {len(codes) - cursor} 个")

    for i in range(count):
        box = inputs.nth(i)
        # 泛微的自定义输入组件在已有值时会把 fill 当成“继续输入”，导致新料号
        # 拼在旧值后面（曾出现两串料号粘连）。先聚焦并全选删除旧内容再写入。
        box.click()
        box.press("Control+A")
        box.press("Delete")
        box.fill(codes[cursor + i])
    log.info("  已填入 %d 个料号", count)
    return cursor + count, "filled"


# ------------------------------------------------------------------ 主流程 --

def run(username: str, password: str) -> int:
    codes = load_material_codes(CODE_FILE)
    log.info("料号清单: %s 共 %d 个", CODE_FILE.name, len(codes))
    if not codes:
        log.error("料号清单为空，退出")
        return 1

    cursor = 0
    stats = {"requests": 0, "filled": 0, "rows": 0, "skipped": 0, "failed": 0}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=HEADLESS)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(PAGE_TIMEOUT_MS)

        completed = False
        try:
            login(page, username, password)
            search_pending_requests(page)

            cells = page.locator(ROW_CELL_SELECTOR)
            count = cells.count()
            log.info("待处理申请共 %d 条", count)
            if count == 0:
                log.warning("没有待处理的料号申请")

            for idx in range(count):
                stats["requests"] += 1
                log.info("处理第 %d/%d 条申请", idx + 1, count)

                # 点击行打开详情：优先新标签页，兼容当前页内跳转/单击无响应
                try:
                    detail, is_new_tab = open_detail(context, page, cells.nth(idx))
                except PWTimeoutError:
                    stats["failed"] += 1
                    log.error("  点击后未打开详情页，跳过")
                    continue

                try:
                    detail.wait_for_load_state("domcontentloaded")
                    rows_before = cursor
                    cursor, state = fill_detail(detail, codes, cursor)
                    if state == "filled":
                        stats["filled"] += 1
                        stats["rows"] += cursor - rows_before
                    elif state == "skipped":
                        stats["skipped"] += 1
                except ValueError as exc:
                    stats["failed"] += 1
                    log.error("  %s，中止后续处理", exc)
                    break
                except PWTimeoutError:
                    stats["failed"] += 1
                    log.error("  详情页加载超时，跳过")
                except Exception:
                    stats["failed"] += 1
                    log.exception("  填写详情页失败，跳过")
                finally:
                    # 新标签页保持打开供人工确认/提交；当前页跳转的需返回列表
                    if not is_new_tab:
                        back_to_list(page)

            log.info("统计: 申请 %d 条，成功填写 %d 条（%d 行），带附件跳过 %d，失败 %d",
                     stats["requests"], stats["filled"], stats["rows"],
                     stats["skipped"], stats["failed"])
            if cursor < len(codes):
                log.warning("料号清单剩余 %d 个未使用", len(codes) - cursor)

            # 执行完成后不自动关闭浏览器：保留登录会话与已填写的详情页，
            # 由人工确认/提交后手动关闭；脚本阻塞到浏览器关闭后自动退出。
            completed = True
            if HEADLESS:
                browser.close()         # 无头模式没有窗口可关，直接释放
            elif browser.is_connected():
                log.info("浏览器保持打开，请人工确认/提交后手动关闭全部浏览器窗口（脚本将自动退出）")
                # 只等 context 的 close 事件不够稳：部分关闭方式下该事件不触发，
                # 而 browser.is_connected() 在浏览器进程结束后仍可能返回 True。
                # 因此每秒带超时地等一次事件，同时轮询页面列表：
                # 所有页面都关闭（或事件/连接报错）即视为浏览器已关闭。
                waited = 0
                while context.pages:
                    try:
                        context.wait_for_event("close", timeout=1000)
                        break
                    except PWTimeoutError:
                        waited += 1
                        if waited % 30 == 0:
                            log.info("  仍在等待：还有 %d 个页面未关闭", len(context.pages))
                    except Exception:
                        break               # 连接断开/浏览器被强制结束
                log.info("浏览器已关闭，脚本退出")
        finally:
            if not completed:           # 中途异常才自动关闭，避免残留进程
                browser.close()
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
