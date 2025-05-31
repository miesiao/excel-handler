import os, re
import pandas as pd
import xlrd
from typing import Union

PHONE_ALIASES = ["電話", "聯絡電話", "電話號碼", "手機", "訂購帳號電話", "收件人電話號碼"]
EMAIL_ALIASES = ["電郵", "Email", "E-mail", "訂購帳號電郵"]
NAME_ALIASES = ["姓名", "訂購帳號姓名"]
PHONE_RE = re.compile(r"\d+")

def read_file(path: str, sheet: Union[int, str] = 0) -> pd.DataFrame:
    if path.lower().endswith(".xlsx"):
        return pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=str)
    elif path.lower().endswith(".xls"):
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(sheet) if isinstance(sheet, int) else wb.sheet_by_name(sheet)
        df = pd.DataFrame([ws.row_values(r) for r in range(ws.nrows)][1:], columns=ws.row_values(0)).astype(str)
    else:
        raise ValueError(f"不支援檔案型別：{path}")
    return df

def norm_phone(raw: str) -> str:
    if pd.isna(raw) or not raw:
        return ""
    digits = "".join(PHONE_RE.findall(str(raw)))
    if digits.startswith("886"):
        digits = digits[3:]
    if digits.startswith("0"):
        digits = digits[1:]
    return "+886" + digits if digits.startswith("9") and len(digits) >= 9 else raw

def is_invalid_row(row: pd.Series) -> bool:
    if any(alias in row.index and "南南" in str(row[alias]) for alias in NAME_ALIASES):
        return True
    for alias in PHONE_ALIASES:
        if alias in row.index:
            digits = "".join(PHONE_RE.findall(str(row[alias])))
            if "09000000" in digits:
                return True
    for alias in EMAIL_ALIASES:
        if alias in row.index and "sousoucorner" in str(row[alias]).lower():
            return True
    return False

def extract_cols(df: pd.DataFrame) -> pd.DataFrame:
    phone_cols, email_cols = [], []
    for col in df.columns:
        if any(a in col for a in PHONE_ALIASES):
            phone_cols.append(col)
        elif any(a in col for a in EMAIL_ALIASES):
            email_cols.append(col)

    phone_cols, email_cols = phone_cols[:2], email_cols[:2]
    rename_map = {c: f"電話{'' if i == 0 else i+1}" for i, c in enumerate(phone_cols)}
    rename_map.update({c: f"電郵{'' if i == 0 else i+1}" for i, c in enumerate(email_cols)})
    df = df.rename(columns=rename_map)[list(rename_map.values())]

    for col in ["電話", "電話2", "電郵", "電郵2"]:
        if col not in df.columns:
            df[col] = ""
    for col in ["電話", "電話2"]:
        df[col] = df[col].apply(norm_phone)
    return df

def process_excel(input_folder, output_path):
    all_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith((".xls", ".xlsx"))]
    if not all_files:
        raise Exception("⚠️ 沒有找到 Excel 檔")

    frames = []
    for path in all_files:
        raw = read_file(path)
        raw = raw[~raw.apply(is_invalid_row, axis=1)]
        frames.append(extract_cols(raw))

    merged = pd.concat(frames, ignore_index=True)
    merged = merged.drop_duplicates(subset=["電話", "電話2", "電郵", "電郵2"], keep="first")
    merged.to_excel(output_path, index=False)
