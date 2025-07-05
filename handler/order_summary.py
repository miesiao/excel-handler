import pandas as pd
import os
import re

# 常數設定
CUSTOMER_COL = '顧客'
NAME_COL = '商品名稱'
PRICE_COL = '商品結帳價'
QTY_COL = '數量'
AMT_COL = '銷售額'
CRAFT_COL = '工藝分類'
SALES_COL = '銷售人員'
TOTAL_KEYS = ['總計', '加總', '合計', 'TOTAL', 'Total']
TIBET_KEYS = ['中國藏區', '中国藏区', '西藏', '藏區', '藏区', '四川藏區', '青海藏區', '甘肅藏區', '雲南藏區']
BRANCH_MAP = {'泰順': '泰順本店', '大安2': '大安2店'}


def load_orders(path):
    if path.endswith(('.xls', '.xlsx')):
        return pd.read_excel(path)
    else:
        return pd.read_csv(path)


def extract_craft(name):
    if pd.isna(name):
        return '未知'
    s = str(name)
    if any(k in s for k in TIBET_KEYS):
        return '北方牧人'
    aaa = re.split(r'[|｜]', s, 1)[0]
    aaa = re.sub(r'【.*?】', '', aaa)
    aaa = aaa.split('・')[0].split('-', 1)[0]
    aaa = re.sub(r'[A-Za-z\\d]+', '', aaa).strip()
    return aaa or '未知'


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if '付款狀態' in df.columns:
        df = df[df['付款狀態'].isin(['已付款', '已部分退款'])]
    df = df[~df[NAME_COL].astype(str).str.contains('|'.join(TOTAL_KEYS), na=False)]
    df = df[~df[CUSTOMER_COL].astype(str).str.contains('南南', na=False)]
    df[CRAFT_COL] = df[NAME_COL].apply(extract_craft)

    amt = []
    for _, row in df.iterrows():
        src = str(row.get('_src', '')).lower()
        if 'pos' in src and '訂單合計' in df:
            a = pd.to_numeric(row.get('訂單合計', 0), errors='coerce')
        elif '付款總金額' in df:
            a = pd.to_numeric(row.get('付款總金額', 0), errors='coerce')
        else:
            a = 0
        amt.append(a)

    df[AMT_COL] = pd.Series(amt).fillna(0)
    return df


def craft_summary(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(CRAFT_COL, as_index=False)[AMT_COL].sum().sort_values(AMT_COL, ascending=False)
    g = pd.concat([g, pd.DataFrame({CRAFT_COL: ['總計'], AMT_COL: [g[AMT_COL].sum()]})], ignore_index=True)
    g[AMT_COL] = g[AMT_COL].apply(lambda x: f"{x:,.0f}")
    return g


def branch_summary(df: pd.DataFrame) -> pd.DataFrame:
    if SALES_COL not in df.columns:
        return pd.DataFrame({'分店': ['資料缺失'], AMT_COL: ['0']})
    df['_b'] = df[SALES_COL].apply(lambda x: next((v for k, v in BRANCH_MAP.items() if k in str(x)), '泰順本店'))
    g = df.groupby('_b', as_index=False)[AMT_COL].sum().sort_values(AMT_COL, ascending=False)
    g = pd.concat([g, pd.DataFrame({'_b': ['總計'], AMT_COL: [g[AMT_COL].sum()]})], ignore_index=True)
    g['_b'] = g['_b'].replace({'其他': '泰順本店'})
    g = g.rename(columns={'_b': '分店'})
    g[AMT_COL] = g[AMT_COL].apply(lambda x: f"{x:,.0f}")
    return g


def add_total(df, label_col):
    total = df[AMT_COL].astype(float).sum()
    return pd.concat([df, pd.DataFrame({label_col: ['總計'], AMT_COL: [total]})], ignore_index=True)


def process_excel(input_folder, output_path):
    all_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith(('.xls', '.xlsx', '.csv'))]
    if not all_files:
        raise Exception("❌ 沒有找到任何 Excel 檔案")

    frames = []
    for f in all_files:
        try:
            df = load_orders(f)
            df['_src'] = os.path.basename(f)
            frames.append(df)
        except Exception as e:
            print(f"❌ 無法讀取檔案 {f}: {e}")

    if not frames:
        raise Exception("⚠️ 所有檔案都讀取失敗")

    raw = pd.concat(frames, ignore_index=True)

    df_online = preprocess(raw[~raw['_src'].str.contains('pos', case=False, na=False)])
    df_pos = preprocess(raw[raw['_src'].str.contains('pos', case=False, na=False)])

    tbl_online = craft_summary(df_online)
    tbl_pos = craft_summary(df_pos)

    # 合併兩表，分類相加
    tbl_all_merge = pd.merge(tbl_online, tbl_pos, how='outer', on='工藝分類', suffixes=('_網店', '_實體店')).fillna(0)
    tbl_all_merge['銷售額'] = pd.to_numeric(tbl_all_merge['銷售額_網店'], errors='coerce') + pd.to_numeric(tbl_all_merge['銷售額_實體店'], errors='coerce')
    tbl_all_merge = tbl_all_merge[['工藝分類', '銷售額']]
    tbl_all = add_total(tbl_all_merge, '工藝分類')
    tbl_all['銷售額'] = tbl_all['銷售額'].apply(lambda x: f"{x:,.0f}")

    tbl_branch = branch_summary(df_pos)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        sheet = 'Summary'
        writer.book.create_sheet(sheet)
        start_col = 0
        col_gap = 2

        def write_block(df_block, title):
            nonlocal start_col
            writer.sheets[sheet].cell(row=1, column=start_col + 1, value=title)
            df_block.to_excel(writer, sheet_name=sheet, index=False, startrow=1, startcol=start_col)
            start_col += df_block.shape[1] + col_gap

        write_block(tbl_all, '1. 網店+實體店')
        write_block(tbl_online, '2. 網店')
        write_block(tbl_pos, '3. 實體店')
        write_block(tbl_branch, '4. 實體店－分店')
