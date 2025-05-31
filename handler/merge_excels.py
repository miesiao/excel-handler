#!/usr/bin/env python
# -*- coding: utf-8 -*-
# v6：合併異格式 Excel → 只留電話/電郵 → 刪「南南」與無效電話/信箱 → 正規化手機 → 去重

import os, glob, re, sys
import pandas as pd
import xlrd                                   # 1.2.0 可讀 .xls
from typing import Union

# -------- 可自行擴充的欄名對應 --------
PHONE_ALIASES = ["電話", "聯絡電話", "電話號碼", "手機", "訂購帳號電話", "收件人電話號碼"]
EMAIL_ALIASES = ["電郵", "Email", "E-mail", "訂購帳號電郵"]
NAME_ALIASES  = ["姓名", "訂購帳號姓名"]

PHONE_RE      = re.compile(r"\d+")

# -------- 讀檔 --------
def read_file(path: str, sheet: Union[int, str] = 0) -> pd.DataFrame:
    if path.lower().endswith(".xlsx"):
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=str)
    elif path.lower().endswith(".xls"):
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(sheet) if isinstance(sheet, int) else wb.sheet_by_name(sheet)
        df = pd.DataFrame([ws.row_values(r) for r in range(ws.nrows)][1:], columns=ws.row_values(0)).astype(str)
    else:
        raise ValueError(f"不支援檔案型別：{path}")
    return df

# -------- 手機正規化 --------
def norm_phone(raw: str) -> str:
    if pd.isna(raw) or not raw:
        return ""
    digits = "".join(PHONE_RE.findall(str(raw)))
    if digits.startswith("886"):
        digits = digits[3:]
    if digits.startswith("0"):
        digits = digits[1:]
    return "+886" + digits if digits.startswith("9") and len(digits) >= 9 else raw

# -------- 判斷列是否需刪除 --------
def is_invalid_row(row: pd.Series) -> bool:
    # 1. 姓名含「南南」
    if any(alias in row.index and "南南" in str(row[alias]) for alias in NAME_ALIASES):
        return True
    # 2. 任何電話欄含 09000000（只看數字）
    for alias in PHONE_ALIASES:
        if alias in row.index:
            digits = "".join(PHONE_RE.findall(str(row[alias])))
            if "09000000" in digits:
                return True
    # 3. 任何 email 含 sousoucorner
    for alias in EMAIL_ALIASES:
        if alias in row.index and "sousoucorner" in str(row[alias]).lower():
            return True
    return False

# -------- 抽取並標準化電話/電郵欄 --------
def extract_cols(df: pd.DataFrame) -> pd.DataFrame:
    phone_cols, email_cols = [], []
    for col in df.columns:
        if any(a in col for a in PHONE_ALIASES):
            phone_cols.append(col)
        elif any(a in col for a in EMAIL_ALIASES):
            email_cols.append(col)

    phone_cols, email_cols = phone_cols[:2], email_cols[:2]        # 最多保留 2 欄
    rename_map = {c: f"電話{'' if i == 0 else i+1}" for i, c in enumerate(phone_cols)}
    rename_map.update({c: f"電郵{'' if i == 0 else i+1}" for i, c in enumerate(email_cols)})
    df = df.rename(columns=rename_map)[list(rename_map.values())]

    # 不足欄位補空
    for col in ["電話", "電話2", "電郵", "電郵2"]:
        if col not in df.columns:
            df[col] = ""

    # 手機號正規化
    for col in ["電話", "電話2"]:
        df[col] = df[col].apply(norm_phone)
    return df

# -------- 主流程 --------
def merge_excels(folder: str, output="merged.xlsx", sheet: Union[int, str] = 0):
    files = sorted(glob.glob(os.path.join(folder, "*.xls*")))
    if not files:
        print("⚠️  沒找到 .xls/.xlsx 檔")
        return

    frames = []
    for f in files:
        raw = read_file(f, sheet)
        # 依新規則批量刪除列
        raw = raw[~raw.apply(is_invalid_row, axis=1)]
        frames.append(extract_cols(raw))

    merged = pd.concat(frames, ignore_index=True)

    # 依「電話×2 + 電郵×2」去重
    merged = merged.drop_duplicates(subset=["電話", "電話2", "電郵", "電郵2"], keep="first")

    merged.to_excel(output, index=False)
    print(f"✅ 已合併 {len(files)} 檔，最終 {len(merged)} 列 → {output}")

# -------- CLI --------
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python merge_excels_v6.py <資料夾> [輸出檔名] [工作表]")
        sys.exit(1)

    folder   = sys.argv[1]
    out_name = sys.argv[2] if len(sys.argv) > 2 else "merged.xlsx"
    sheet_no = sys.argv[3] if len(sys.argv) > 3 else 0
    merge_excels(folder, out_name, sheet_no)
