from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Any

import pandas as pd
import requests


COMPANY_API_URL = (
    "https://data.gcis.nat.gov.tw/od/data/api/"
    "5F64D864-61CB-4D0D-8AD9-492047CC1EA6"
)

INPUT_FILE = Path("company_list_fbfh_enriched.xlsx")
OUTPUT_FILE = Path("company_list_fbfh_enriched_api.xlsx")
MAIN_SHEET = "公司基本資料"

REQUEST_DELAY = 0.5
TIMEOUT_SECONDS = 30
MAX_RETRIES = 3
SAVE_EVERY = 25

FORCE_RERUN_SUCCESS = False
FILL_MISSING_CAPITAL = True
CAPITAL_FILL_FIELD = "Capital_Stock_Amount"

API_COLUMNS = [
    "商工API查詢統編",
    "商工API查詢狀態",
    "商工API錯誤訊息",
    "商工公司名稱",
    "商工負責人",
    "商工公司狀態",
    "商工資本總額",
    "商工實收資本額",
    "商工資本總額(萬元)",
    "商工實收資本額(萬元)",
    "商工公司地址",
    "商工核准設立日期",
    "商工最後核准變更日期",
    "資本額補值狀態",
    "資本額資料來源",
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
    return "" if text.lower() == "nan" else text


def normalize_tax_id(value: Any) -> str:
    text = clean_text(value)
    if text.endswith(".0"):
        text = text[:-2]
    digits = "".join(ch for ch in text if ch.isdigit())
    return digits.zfill(8) if digits else ""


def query_company(session: requests.Session, tax_id: str):
    params = {
        "$format": "json",
        "$filter": f"Business_Accounting_NO eq {tax_id}",
        "$skip": 0,
        "$top": 50,
    }

    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(
                COMPANY_API_URL,
                params=params,
                timeout=TIMEOUT_SECONDS,
            )
            response.raise_for_status()

            result = response.json()

            if not isinstance(result, list):
                return None, f"API 格式異常：{type(result).__name__}"

            if not result:
                return None, "查無資料"

            return result[0], ""

        except requests.exceptions.JSONDecodeError:
            preview = response.text[:200].strip()
            last_error = f"JSON 解析失敗：{preview or '空白內容'}"

        except requests.RequestException as exc:
            last_error = f"HTTP 錯誤：{exc}"

        except ValueError as exc:
            last_error = f"資料解析失敗：{exc}"

        if attempt < MAX_RETRIES:
            wait_seconds = attempt * 2
            print(
                f"    第 {attempt} 次失敗，"
                f"{wait_seconds} 秒後重試：{last_error}"
            )
            time.sleep(wait_seconds)

    return None, last_error


def fill_api_fields(
    df: pd.DataFrame,
    idx: int,
    tax_id: str,
    company: dict[str, Any] | None,
    error_message: str,
) -> None:
    df.at[idx, "商工API查詢統編"] = tax_id

    if company is None:
        df.at[idx, "商工API查詢狀態"] = (
            "查無資料" if error_message == "查無資料" else "查詢失敗"
        )
        df.at[idx, "商工API錯誤訊息"] = error_message
        return

    capital = company.get("Capital_Stock_Amount")
    paid = company.get("Paid_In_Capital_Amount")

    df.at[idx, "商工API查詢狀態"] = "成功"
    df.at[idx, "商工API錯誤訊息"] = ""
    df.at[idx, "商工公司名稱"] = clean_text(company.get("Company_Name", ""))
    df.at[idx, "商工負責人"] = clean_text(company.get("Responsible_Name", ""))
    df.at[idx, "商工公司狀態"] = clean_text(company.get("Company_Status_Desc", ""))
    df.at[idx, "商工資本總額"] = clean_text(capital)
    df.at[idx, "商工實收資本額"] = clean_text(paid)
    df.at[idx, "商工公司地址"] = clean_text(company.get("Company_Location", ""))
    df.at[idx, "商工核准設立日期"] = clean_text(company.get("Company_Setup_Date", ""))
    df.at[idx, "商工最後核准變更日期"] = clean_text(
        company.get("Change_Of_Approval_Data", "")
    )

    capital_num = pd.to_numeric(capital, errors="coerce")
    paid_num = pd.to_numeric(paid, errors="coerce")

    if pd.notna(capital_num):
        df.at[idx, "商工資本總額(萬元)"] = str(capital_num / 10000)

    if pd.notna(paid_num):
        df.at[idx, "商工實收資本額(萬元)"] = str(paid_num / 10000)

    if FILL_MISSING_CAPITAL and "資本額" in df.columns:
        current = clean_text(df.at[idx, "資本額"])

        if current == "":
            fill_value = company.get(CAPITAL_FILL_FIELD)
            fill_num = pd.to_numeric(fill_value, errors="coerce")

            if pd.notna(fill_num):
                df.at[idx, "資本額"] = str(int(fill_num))
                df.at[idx, "資本額補值狀態"] = "已補值"
                df.at[idx, "資本額資料來源"] = "商工API資本總額（元）"
            else:
                df.at[idx, "資本額補值狀態"] = "API無可用資本額"
        else:
            df.at[idx, "資本額補值狀態"] = "原本已有資本額"


def save_all_sheets(sheets: dict[str, pd.DataFrame]) -> None:
    temp = OUTPUT_FILE.with_name(
        f"{OUTPUT_FILE.stem}.tmp{OUTPUT_FILE.suffix}"
    )

    with pd.ExcelWriter(temp, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
            )

    os.replace(temp, OUTPUT_FILE)


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"找不到輸入檔：{INPUT_FILE.resolve()}"
        )

    sheets = pd.read_excel(
        INPUT_FILE,
        sheet_name=None,
        dtype=str,
    )

    if MAIN_SHEET not in sheets:
        raise KeyError(f"找不到工作表：{MAIN_SHEET}")

    df = sheets[MAIN_SHEET].fillna("")

    if "統一編號" not in df.columns:
        raise KeyError(
            "主工作表缺少「統一編號」欄。"
            "請先執行 FBFH 通用版取得統編。"
        )

    for col in API_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    target_indices = []

    for idx, row in df.iterrows():
        tax_id = normalize_tax_id(row.get("統一編號", ""))

        if not tax_id:
            continue

        if (
            not FORCE_RERUN_SUCCESS
            and clean_text(row.get("商工API查詢狀態", "")) == "成功"
        ):
            continue

        target_indices.append(idx)

    print("=" * 68)
    print("商工 API 通用補齊版")
    print(f"主表筆數：{len(df)}")
    print(f"本次需查：{len(target_indices)}")
    print("=" * 68)

    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "Python company research"
        )
    })

    cache = {}

    try:
        for sequence, idx in enumerate(target_indices, start=1):
            tax_id = normalize_tax_id(df.at[idx, "統一編號"])
            company_name = (
                clean_text(df.at[idx, "公司名稱"])
                if "公司名稱" in df.columns
                else ""
            )

            print(
                f"[{sequence}/{len(target_indices)}] "
                f"{company_name}｜{tax_id}"
            )

            if tax_id in cache:
                company, error_message = cache[tax_id]
                print("    使用已查詢快取")
            else:
                company, error_message = query_company(session, tax_id)
                cache[tax_id] = (company, error_message)
                time.sleep(REQUEST_DELAY)

            fill_api_fields(
                df=df,
                idx=idx,
                tax_id=tax_id,
                company=company,
                error_message=error_message,
            )

            print(
                f"    狀態：{df.at[idx, '商工API查詢狀態']}"
                f"｜公司：{df.at[idx, '商工公司名稱']}"
                f"｜資本總額：{df.at[idx, '商工資本總額']}"
            )

            if (
                sequence % SAVE_EVERY == 0
                or sequence == len(target_indices)
            ):
                sheets[MAIN_SHEET] = df
                save_all_sheets(sheets)
                print(f"    已儲存：{OUTPUT_FILE}")

    except KeyboardInterrupt:
        print("\n偵測到中止，正在保存目前成果……")

    finally:
        session.close()
        sheets[MAIN_SHEET] = df
        save_all_sheets(sheets)

    counts = (
        df["商工API查詢狀態"]
        .fillna("")
        .value_counts()
        .to_dict()
    )

    print("\n" + "=" * 68)
    print("完成")
    print(f"成功：{int(counts.get('成功', 0))}")
    print(f"查無資料：{int(counts.get('查無資料', 0))}")
    print(f"查詢失敗：{int(counts.get('查詢失敗', 0))}")
    print(f"輸出：{OUTPUT_FILE}")
    print("=" * 68)


if __name__ == "__main__":
    main()
