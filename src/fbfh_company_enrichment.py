from __future__ import annotations

import os
import re
import time
import tkinter as tk
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ============================================================
# FBFH 純公司名稱通用版
#
# 最低需求：
# - Excel/CSV 中有一欄「公司名稱」
#
# 功能：
# - 支援 xlsx / xls / csv
# - 保留原始所有欄位
# - 自動新增 FBFH 基本資料
# - 抓進出口資格與歷年進出口級距
# - 驗證碼小視窗
# - 自動捲到驗證碼位置
# - 每家公司完成立即存檔
# - SUCCESS 自動跳過，可中斷續跑
# ============================================================

INPUT_FILE = Path("company_list.xlsx")
INPUT_SHEET: str | None = None   # None = 第一張工作表
COMPANY_COLUMN = "公司名稱"

OUTPUT_FILE = Path("company_list_fbfh_enriched.xlsx")
OUTPUT_SHEET = "公司基本資料"
HISTORY_SHEET = "歷年實績級距"
SUMMARY_SHEET = "執行摘要"

URL = "https://fbfh.trade.gov.tw/fb/web/queryBasicf.do"

LIMIT: int | None = None
FORCE_RERUN_SUCCESS = False

ERROR_SCREENSHOT_DIR = Path("fbfh_company_error_screenshots")
PROGRESS_REPORT_EVERY = 10


FBFH_COLUMNS = [
    "統一編號",
    "FBFH_中文名稱",
    "FBFH_英文名稱",
    "FBFH_負責人",
    "FBFH_地址",
    "FBFH_英文地址",
    "FBFH_電話",
    "FBFH_電話2",
    "FBFH_傳真",
    "FBFH_Email",
    "FBFH_網址",
    "FBFH_進口產品",
    "FBFH_出口產品",
    "進口資格",
    "出口資格",
    "最新級距期間",
    "最新進口級距",
    "最新出口級距",
    "最近完整年度期間",
    "最近完整年度進口級距",
    "最近完整年度出口級距",
    "查詢狀態",
    "配對方式",
    "查詢時間",
    "錯誤訊息",
    "錯誤截圖",
]


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    return "" if text.lower() == "nan" else re.sub(r"\s+", " ", text).strip()


def normalize_company_name(value: Any) -> str:
    return (
        unicodedata.normalize("NFKC", clean_text(value))
        .replace(" ", "")
        .replace("　", "")
        .replace("臺", "台")
    )


def safe_filename(value: str, max_length: int = 60) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', "_", clean_text(value))
    return value[:max_length] or "unknown"


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def read_input_file() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"找不到輸入檔：{INPUT_FILE.resolve()}")

    suffix = INPUT_FILE.suffix.lower()

    if suffix in {".xlsx", ".xls"}:
        df = pd.read_excel(
            INPUT_FILE,
            sheet_name=0 if INPUT_SHEET is None else INPUT_SHEET,
            dtype=str,
        )
    elif suffix == ".csv":
        try:
            df = pd.read_csv(INPUT_FILE, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(INPUT_FILE, dtype=str, encoding="big5")
    else:
        raise ValueError("目前只支援 .xlsx / .xls / .csv")

    df = df.fillna("")

    if COMPANY_COLUMN not in df.columns:
        raise KeyError(
            f"找不到『{COMPANY_COLUMN}』欄位，目前欄位：{list(df.columns)}"
        )

    df[COMPANY_COLUMN] = df[COMPANY_COLUMN].astype(str).str.strip()
    df = df[df[COMPANY_COLUMN] != ""].reset_index(drop=True)

    if LIMIT is not None:
        df = df.iloc[:LIMIT].copy().reset_index(drop=True)

    return df


def prepare_main_df() -> pd.DataFrame:
    df = read_input_file()

    for col in FBFH_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    mask = df["查詢狀態"].astype(str).str.strip().eq("")
    df.loc[mask, "查詢狀態"] = "PENDING"

    return df


def load_resume_data(base_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    history_columns = [
        "公司名稱",
        "統一編號",
        "年度期間",
        "進口級距",
        "出口級距",
    ]

    if not OUTPUT_FILE.exists():
        return base_df, pd.DataFrame(columns=history_columns)

    old = pd.read_excel(
        OUTPUT_FILE,
        sheet_name=OUTPUT_SHEET,
        dtype=str,
    ).fillna("")

    try:
        history = pd.read_excel(
            OUTPUT_FILE,
            sheet_name=HISTORY_SHEET,
            dtype=str,
        ).fillna("")
    except ValueError:
        history = pd.DataFrame(columns=history_columns)

    old_map = {}
    for _, row in old.iterrows():
        key = normalize_company_name(row.get(COMPANY_COLUMN, ""))
        if key:
            old_map[key] = row

    for idx, row in base_df.iterrows():
        key = normalize_company_name(row[COMPANY_COLUMN])
        old_row = old_map.get(key)
        if old_row is None:
            continue
        for col in FBFH_COLUMNS:
            if col in old_row.index:
                base_df.at[idx, col] = clean_text(old_row[col])

    return base_df, history


def build_summary(main_df: pd.DataFrame, history_df: pd.DataFrame) -> pd.DataFrame:
    counts = main_df["查詢狀態"].fillna("").value_counts().to_dict()

    return pd.DataFrame([
        {"項目": "輸出更新時間", "數值": now_text()},
        {"項目": "名單總筆數", "數值": len(main_df)},
        {"項目": "SUCCESS", "數值": counts.get("SUCCESS", 0)},
        {"項目": "NOT_FOUND", "數值": counts.get("NOT_FOUND", 0)},
        {"項目": "REVIEW", "數值": counts.get("REVIEW", 0)},
        {"項目": "SKIPPED", "數值": counts.get("SKIPPED", 0)},
        {"項目": "ERROR", "數值": counts.get("ERROR", 0)},
        {"項目": "PENDING", "數值": counts.get("PENDING", 0)},
        {"項目": "歷年實績資料列數", "數值": len(history_df)},
    ])


def format_workbook(path: Path) -> None:
    wb = load_workbook(path)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    widths = {
        COMPANY_COLUMN: 34,
        "統一編號": 14,
        "FBFH_中文名稱": 34,
        "FBFH_英文名稱": 38,
        "FBFH_負責人": 14,
        "FBFH_地址": 42,
        "FBFH_英文地址": 48,
        "FBFH_電話": 18,
        "FBFH_電話2": 18,
        "FBFH_傳真": 18,
        "FBFH_Email": 36,
        "FBFH_網址": 36,
        "FBFH_進口產品": 36,
        "FBFH_出口產品": 36,
        "最新級距期間": 28,
        "最近完整年度期間": 28,
        "錯誤訊息": 45,
        "錯誤截圖": 42,
        "年度期間": 28,
    }

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        if ws.max_row >= 1:
            ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        headers = {
            clean_text(ws.cell(1, c).value): c
            for c in range(1, ws.max_column + 1)
        }

        for c in range(1, ws.max_column + 1):
            header = clean_text(ws.cell(1, c).value)
            ws.column_dimensions[get_column_letter(c)].width = widths.get(header, 16)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        tax_col = headers.get("統一編號")
        if tax_col:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, tax_col).number_format = "@"

    wb.save(path)


def atomic_save_excel(main_df: pd.DataFrame, history_df: pd.DataFrame) -> None:
    temp = OUTPUT_FILE.with_name(
        f"{OUTPUT_FILE.stem}.tmp{OUTPUT_FILE.suffix}"
    )

    with pd.ExcelWriter(temp, engine="openpyxl") as writer:
        main_df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
        history_df.to_excel(writer, sheet_name=HISTORY_SHEET, index=False)
        build_summary(main_df, history_df).to_excel(
            writer,
            sheet_name=SUMMARY_SHEET,
            index=False,
        )

    format_workbook(temp)
    os.replace(temp, OUTPUT_FILE)


def find_first(driver, candidates, timeout=8, clickable=False):
    condition = EC.element_to_be_clickable if clickable else EC.visibility_of_element_located

    for by, selector in candidates:
        try:
            return WebDriverWait(driver, timeout).until(condition((by, selector)))
        except TimeoutException:
            continue

    raise NoSuchElementException(f"找不到元素：{candidates}")


def find_company_input(driver):
    return find_first(driver, [
        (By.XPATH, "//input[contains(@placeholder,'至少輸入二碼中文名稱')]"),
        (By.XPATH, "//*[contains(normalize-space(.),'廠商中文名稱')]/following::input[1]"),
    ])


def find_captcha_input(driver):
    return find_first(driver, [
        (By.XPATH, "//input[contains(@placeholder,'點擊上圖可重新產生驗證碼')]"),
        (By.XPATH, "//*[contains(normalize-space(.),'驗證碼')]/following::input[1]"),
    ])


def find_search_button(driver):
    return find_first(driver, [
        (By.XPATH, "//button[contains(normalize-space(.),'查詢')]"),
        (By.XPATH, "//*[self::button or self::a][contains(normalize-space(.),'Search') and contains(normalize-space(.),'查詢')]"),
    ], clickable=True)


def scroll_captcha_to_center(driver) -> None:
    captcha_input = find_captcha_input(driver)
    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            behavior: 'instant',
            block: 'center',
            inline: 'nearest'
        });
        """,
        captcha_input,
    )
    driver.execute_script("window.scrollBy(0, -110);")
    time.sleep(0.4)


def click_refresh_captcha(driver) -> None:
    candidates = [
        (By.XPATH, "//img[contains(@alt,'驗證碼') or contains(@title,'驗證碼')]"),
        (By.XPATH, "//*[contains(normalize-space(.),'點擊圖片可重新產生驗證碼')]/preceding::img[1]"),
    ]

    for by, selector in candidates:
        elements = driver.find_elements(by, selector)
        if elements:
            try:
                driver.execute_script("arguments[0].click();", elements[0])
                time.sleep(0.6)
                return
            except Exception:
                pass


def classify_search_page(driver) -> str:
    source = driver.page_source
    body = clean_text(driver.find_element(By.TAG_NAME, "body").text)

    if "符合查詢結果共" in source or "fuzzy search results total" in source:
        return "RESULT"

    if any(k in body for k in ["查無資料", "查無符合", "無符合查詢", "total 0 items", "共0筆"]):
        return "NOT_FOUND"

    if any(k in body for k in ["驗證碼錯誤", "驗證碼不正確", "請輸入正確驗證碼", "Verification code error"]):
        return "CAPTCHA_ERROR"

    return "UNKNOWN"


def wait_for_search_outcome(driver, timeout=18) -> str:
    try:
        return WebDriverWait(driver, timeout).until(
            lambda d: (
                result
                if (result := classify_search_page(d)) != "UNKNOWN"
                else False
            )
        )
    except TimeoutException:
        return classify_search_page(driver)


def get_result_candidates(driver) -> list[dict]:
    candidates = []

    for row in driver.find_elements(By.XPATH, "//table//tr[td]"):
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) < 4:
            continue

        texts = [clean_text(cell.text) for cell in cells]

        if not texts or not re.fullmatch(r"\d{8}", texts[0]):
            continue

        candidates.append({
            "統一編號": texts[0],
            "中文名稱": texts[1] if len(texts) > 1 else "",
            "英文名稱": texts[2] if len(texts) > 2 else "",
            "負責人": texts[3] if len(texts) > 3 else "",
        })

    return candidates


def select_result_candidate(candidates: list[dict], original_name: str):
    if not candidates:
        return "NOT_FOUND", None

    wanted = normalize_company_name(original_name)

    exact = [
        row for row in candidates
        if normalize_company_name(row["中文名稱"]) == wanted
    ]

    if len(exact) == 1:
        return "EXACT", exact[0]

    if len(candidates) == 1:
        return "SINGLE_FUZZY", candidates[0]

    return "MULTIPLE", None


def close_modal(driver, modal_id: str) -> None:
    driver.execute_script(
        """
        const modal = document.querySelector(arguments[0]);
        if (!modal) return;

        if (window.jQuery) {
            window.jQuery(modal).modal('hide');
        } else {
            modal.classList.remove('show');
            modal.style.display = 'none';
        }

        document.body.classList.remove('modal-open');
        document.querySelectorAll('.modal-backdrop').forEach(x => x.remove());
        """,
        f"#{modal_id}",
    )

    try:
        WebDriverWait(driver, 6).until(
            EC.invisibility_of_element_located((By.ID, modal_id))
        )
    except TimeoutException:
        pass


def get_basic_fields(driver, business_id: str) -> dict:
    driver.execute_script("kdbase_showPopBasic(arguments[0]);", business_id)

    modal = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.ID, "kdbase_popBasic"))
    )

    WebDriverWait(driver, 15).until(
        lambda d: business_id in clean_text(
            d.find_element(By.ID, "popBasicCardBody").text
        )
    )

    result = {
        "中文名稱": "",
        "英文名稱": "",
        "負責人": "",
        "地址": "",
        "英文地址": "",
        "電話": "",
        "電話2": "",
        "傳真": "",
        "Email": "",
        "網址": "",
        "進口產品": "",
        "出口產品": "",
        "進口資格": "",
        "出口資格": "",
    }

    for row in modal.find_elements(
        By.XPATH,
        ".//div[@id='popBasicCardBody']//table//tr",
    ):
        cells = row.find_elements(By.XPATH, "./th|./td")
        if len(cells) < 2:
            continue

        label = clean_text(cells[0].text)
        value = clean_text(cells[1].text)

        if "廠商中文名稱" in label:
            result["中文名稱"] = value
        elif "廠商英文名稱" in label:
            result["英文名稱"] = value
        elif "中文營業地址" in label:
            result["地址"] = value
        elif "英文營業地址" in label:
            result["英文地址"] = value
        elif "代表人" in label:
            result["負責人"] = value
        elif "電話號碼 1" in label:
            result["電話"] = value
        elif "電話號碼 2" in label:
            result["電話2"] = value
        elif "傳真號碼" in label:
            result["傳真"] = value
        elif "電子信箱" in label:
            result["Email"] = value
        elif "廠商網址" in label:
            result["網址"] = value
        elif "產品項目進口" in label:
            result["進口產品"] = value
        elif "產品項目出口" in label:
            result["出口產品"] = value
        elif "進口資格" in label:
            result["進口資格"] = value
        elif "出口資格" in label:
            result["出口資格"] = value

    close_modal(driver, "kdbase_popBasic")
    return result


def open_and_parse_grades(driver, business_id: str, company_name: str) -> list[dict]:
    driver.execute_script("kdbase_showPopGrade(arguments[0]);", business_id)

    WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.ID, "kdbase_popGrade"))
    )

    WebDriverWait(driver, 20).until(
        lambda d: clean_text(d.find_element(By.ID, "banNoM2").text) == business_id
    )

    WebDriverWait(driver, 20).until(
        lambda d: len(
            d.find_elements(By.CSS_SELECTOR, "#kdbase_popGradeList tr")
        ) > 0
    )

    rows_out = []

    for row in driver.find_elements(
        By.CSS_SELECTOR,
        "#kdbase_popGradeList tr",
    ):
        cells = row.find_elements(By.XPATH, "./th|./td")
        values = [clean_text(cell.text) for cell in cells]

        if len(values) >= 3 and re.search(r"\d{3,4}", values[0]):
            rows_out.append({
                "公司名稱": company_name,
                "統一編號": business_id,
                "年度期間": values[0],
                "進口級距": values[1],
                "出口級距": values[2],
            })

    close_modal(driver, "kdbase_popGrade")
    return rows_out


def summarize_grades(grade_rows: list[dict]) -> dict:
    result = {
        "最新級距期間": "",
        "最新進口級距": "",
        "最新出口級距": "",
        "最近完整年度期間": "",
        "最近完整年度進口級距": "",
        "最近完整年度出口級距": "",
    }

    if not grade_rows:
        return result

    latest = grade_rows[0]

    result["最新級距期間"] = latest["年度期間"]
    result["最新進口級距"] = latest["進口級距"]
    result["最新出口級距"] = latest["出口級距"]

    for row in grade_rows:
        period = row["年度期間"]

        if "01月至12月" in period or "Jan. - Dec." in period or "Jan.-Dec." in period:
            result["最近完整年度期間"] = period
            result["最近完整年度進口級距"] = row["進口級距"]
            result["最近完整年度出口級距"] = row["出口級距"]
            break

    return result


def show_captcha_dialog(
    company_name: str,
    sequence: int,
    total: int,
    status_message: str = "",
):
    result = {"action": "QUIT", "value": ""}

    root = tk.Tk()
    root.title("驗證碼輸入")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    width = 430
    height = 230

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x = max(20, screen_width - width - 40)
    y = max(20, screen_height - height - 90)

    root.geometry(f"{width}x{height}+{x}+{y}")

    outer = tk.Frame(root, padx=16, pady=14)
    outer.pack(fill="both", expand=True)

    tk.Label(
        outer,
        text=f"進度：{sequence} / {total}",
        font=("Microsoft JhengHei UI", 10, "bold"),
        anchor="w",
    ).pack(fill="x")

    tk.Label(
        outer,
        text=company_name,
        font=("Microsoft JhengHei UI", 11),
        anchor="w",
        justify="left",
        wraplength=395,
    ).pack(fill="x", pady=(5, 8))

    if status_message:
        tk.Label(
            outer,
            text=status_message,
            font=("Microsoft JhengHei UI", 9),
            fg="#A00000",
            anchor="w",
            justify="left",
            wraplength=395,
        ).pack(fill="x", pady=(0, 7))

    entry = tk.Entry(
        outer,
        font=("Consolas", 18),
        justify="center",
    )
    entry.pack(fill="x", ipady=5)
    entry.focus_force()

    tk.Label(
        outer,
        text="看瀏覽器中的驗證碼後直接輸入；按 Enter 送出",
        font=("Microsoft JhengHei UI", 9),
        anchor="w",
    ).pack(fill="x", pady=(5, 10))

    buttons = tk.Frame(outer)
    buttons.pack(fill="x")

    def finish(action: str, value: str = ""):
        result["action"] = action
        result["value"] = value.strip()
        root.quit()
        root.destroy()

    def submit():
        value = entry.get().strip()
        if not value:
            root.bell()
            entry.focus_force()
            return
        finish("SUBMIT", value)

    tk.Button(buttons, text="送出", width=9, command=submit).pack(side="left")
    tk.Button(
        buttons,
        text="更新驗證碼",
        width=11,
        command=lambda: finish("REFRESH"),
    ).pack(side="left", padx=(8, 0))
    tk.Button(
        buttons,
        text="略過",
        width=8,
        command=lambda: finish("SKIP"),
    ).pack(side="left", padx=(8, 0))
    tk.Button(
        buttons,
        text="儲存並結束",
        width=11,
        command=lambda: finish("QUIT"),
    ).pack(side="right")

    root.bind("<Return>", lambda _e: submit())
    root.bind("<Escape>", lambda _e: finish("QUIT"))
    root.protocol("WM_DELETE_WINDOW", lambda: finish("QUIT"))

    root.after(120, root.lift)
    root.after(160, lambda: root.attributes("-topmost", True))
    root.after(200, entry.focus_force)

    root.mainloop()

    return result["action"], result["value"]


def save_error_screenshot(driver, sequence: int, company_name: str, suffix: str) -> str:
    ERROR_SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

    path = ERROR_SCREENSHOT_DIR / (
        f"{sequence:04d}_{safe_filename(company_name)}_{suffix}.png"
    )

    try:
        driver.save_screenshot(str(path))
        return str(path)
    except Exception:
        return ""


def replace_history_rows(
    history_df: pd.DataFrame,
    company_name: str,
    business_id: str,
    new_rows: list[dict],
) -> pd.DataFrame:
    columns = [
        "公司名稱",
        "統一編號",
        "年度期間",
        "進口級距",
        "出口級距",
    ]

    if history_df.empty:
        base = pd.DataFrame(columns=columns)
    else:
        base = history_df.copy()

    company_key = normalize_company_name(company_name)

    keep = []
    for _, row in base.iterrows():
        same_id = clean_text(row.get("統一編號", "")) == business_id
        same_name = normalize_company_name(row.get("公司名稱", "")) == company_key
        keep.append(not (same_id or same_name))

    if len(base) > 0:
        base = base.loc[keep].copy()

    if new_rows:
        base = pd.concat([base, pd.DataFrame(new_rows)], ignore_index=True)

    return base[columns]


def query_one_company(
    driver,
    company_name: str,
    sequence: int,
    total: int,
):
    update = {col: "" for col in FBFH_COLUMNS}
    update["查詢狀態"] = "PENDING"
    update["查詢時間"] = now_text()

    driver.get(URL)

    WebDriverWait(driver, 15).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    company_input = find_company_input(driver)
    company_input.clear()
    company_input.send_keys(company_name)

    dialog_message = ""
    scroll_captcha_to_center(driver)

    while True:
        action, captcha = show_captcha_dialog(
            company_name,
            sequence,
            total,
            dialog_message,
        )

        dialog_message = ""

        if action == "REFRESH":
            click_refresh_captcha(driver)
            scroll_captcha_to_center(driver)
            continue

        if action == "SKIP":
            update["查詢狀態"] = "SKIPPED"
            update["錯誤訊息"] = "使用者手動略過"
            update["查詢時間"] = now_text()
            return update, [], "CONTINUE"

        if action == "QUIT":
            return update, [], "QUIT"

        captcha_input = find_captcha_input(driver)
        captcha_input.clear()
        captcha_input.send_keys(captcha)

        find_search_button(driver).click()

        outcome = wait_for_search_outcome(driver)

        if outcome in {"CAPTCHA_ERROR", "UNKNOWN"}:
            dialog_message = "驗證碼可能錯誤或查詢未完成，請查看新的驗證碼後再輸入。"

            try:
                company_input = find_company_input(driver)
                if not clean_text(company_input.get_attribute("value")):
                    company_input.send_keys(company_name)
            except Exception:
                driver.get(URL)
                company_input = find_company_input(driver)
                company_input.clear()
                company_input.send_keys(company_name)

            scroll_captcha_to_center(driver)
            continue

        if outcome == "NOT_FOUND":
            update["查詢狀態"] = "NOT_FOUND"
            update["錯誤訊息"] = "進出口廠商登記系統查無資料"
            update["錯誤截圖"] = save_error_screenshot(
                driver,
                sequence,
                company_name,
                "NOT_FOUND",
            )
            update["查詢時間"] = now_text()
            return update, [], "CONTINUE"

        candidates = get_result_candidates(driver)
        match_type, selected = select_result_candidate(candidates, company_name)

        if match_type == "MULTIPLE":
            update["查詢狀態"] = "REVIEW"
            update["配對方式"] = "MULTIPLE_RESULTS"
            update["錯誤訊息"] = (
                "查詢結果有多筆，需人工確認："
                + " | ".join(
                    f"{row['統一編號']} {row['中文名稱']}"
                    for row in candidates
                )
            )
            update["錯誤截圖"] = save_error_screenshot(
                driver,
                sequence,
                company_name,
                "MULTIPLE",
            )
            update["查詢時間"] = now_text()
            return update, [], "CONTINUE"

        if not selected:
            update["查詢狀態"] = "ERROR"
            update["錯誤訊息"] = "結果頁存在，但無法辨識公司資料列"
            update["錯誤截圖"] = save_error_screenshot(
                driver,
                sequence,
                company_name,
                "NO_ROW",
            )
            update["查詢時間"] = now_text()
            return update, [], "CONTINUE"

        business_id = selected["統一編號"]
        basic = get_basic_fields(driver, business_id)

        update.update({
            "統一編號": business_id,
            "FBFH_中文名稱": basic["中文名稱"],
            "FBFH_英文名稱": basic["英文名稱"],
            "FBFH_負責人": basic["負責人"],
            "FBFH_地址": basic["地址"],
            "FBFH_英文地址": basic["英文地址"],
            "FBFH_電話": basic["電話"],
            "FBFH_電話2": basic["電話2"],
            "FBFH_傳真": basic["傳真"],
            "FBFH_Email": basic["Email"],
            "FBFH_網址": basic["網址"],
            "FBFH_進口產品": basic["進口產品"],
            "FBFH_出口產品": basic["出口產品"],
            "進口資格": basic["進口資格"],
            "出口資格": basic["出口資格"],
            "配對方式": match_type,
        })

        grade_rows = open_and_parse_grades(
            driver,
            business_id,
            company_name,
        )

        update.update(summarize_grades(grade_rows))

        update["查詢狀態"] = "SUCCESS"
        update["查詢時間"] = now_text()

        return update, grade_rows, "CONTINUE"


def main():
    main_df = prepare_main_df()
    main_df, history_df = load_resume_data(main_df)

    total = len(main_df)
    completed = int((main_df["查詢狀態"] == "SUCCESS").sum())

    print("=" * 76)
    print("FBFH 純公司名稱通用版")
    print(f"輸入：{INPUT_FILE.resolve()}")
    print(f"公司名稱欄：{COMPANY_COLUMN}")
    print(f"名單筆數：{total}")
    print(f"既有 SUCCESS：{completed}")
    print("只要有『公司名稱』欄即可，其餘原始欄位會全部保留。")
    print("=" * 76)

    atomic_save_excel(main_df, history_df)

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)

    processed_this_run = 0
    should_quit = False

    try:
        for idx in range(total):
            company_name = clean_text(main_df.at[idx, COMPANY_COLUMN])
            prior_status = clean_text(main_df.at[idx, "查詢狀態"])
            sequence = idx + 1

            if prior_status == "SUCCESS" and not FORCE_RERUN_SUCCESS:
                print(f"[{sequence}/{total}] 已完成，跳過：{company_name}")
                continue

            try:
                update, grade_rows, action = query_one_company(
                    driver,
                    company_name,
                    sequence,
                    total,
                )

                if action == "QUIT":
                    print("收到儲存並結束指令。")
                    should_quit = True
                    break

                for col in FBFH_COLUMNS:
                    if col in update:
                        main_df.at[idx, col] = clean_text(update[col])

                if (
                    update.get("查詢狀態") == "SUCCESS"
                    and clean_text(update.get("統一編號"))
                ):
                    history_df = replace_history_rows(
                        history_df,
                        company_name,
                        clean_text(update["統一編號"]),
                        grade_rows,
                    )

                atomic_save_excel(main_df, history_df)

                processed_this_run += 1

                print(
                    f"狀態：{main_df.at[idx, '查詢狀態']}"
                    f"｜已即時儲存：{OUTPUT_FILE.name}"
                )

            except KeyboardInterrupt:
                print("\n偵測到 Ctrl+C，正在保存目前成果……")
                should_quit = True
                break

            except Exception as exc:
                main_df.at[idx, "查詢狀態"] = "ERROR"
                main_df.at[idx, "查詢時間"] = now_text()
                main_df.at[idx, "錯誤訊息"] = f"{type(exc).__name__}: {exc}"
                main_df.at[idx, "錯誤截圖"] = save_error_screenshot(
                    driver,
                    sequence,
                    company_name,
                    "ERROR",
                )

                atomic_save_excel(main_df, history_df)
                processed_this_run += 1

                print(
                    f"本筆 ERROR，已保存後繼續："
                    f"{main_df.at[idx, '錯誤訊息']}"
                )

            if (
                processed_this_run > 0
                and processed_this_run % PROGRESS_REPORT_EVERY == 0
            ):
                success_count = int(
                    (main_df["查詢狀態"] == "SUCCESS").sum()
                )

                print("=" * 76)
                print(
                    f"本次已處理 {processed_this_run} 筆；"
                    f"累計 SUCCESS {success_count}/{total}"
                )
                print("=" * 76)

    finally:
        atomic_save_excel(main_df, history_df)
        driver.quit()

    counts = main_df["查詢狀態"].value_counts().to_dict()

    print("=" * 76)
    print("程式已結束，成果已保存。")
    print(f"輸出檔案：{OUTPUT_FILE.resolve()}")

    for status in [
        "SUCCESS",
        "NOT_FOUND",
        "REVIEW",
        "SKIPPED",
        "ERROR",
        "PENDING",
    ]:
        print(f"{status}：{int(counts.get(status, 0))}")

    print(f"歷年實績：{len(history_df)} 筆")
    print("=" * 76)

    if should_quit:
        print("下次直接重新執行同一支程式，即會自動跳過 SUCCESS 並續跑。")


if __name__ == "__main__":
    main()
