"""
generate_schedule_fixed_rev9.py

Revision History:
- Rev1: 初期スクリプト作成（ヘッダー行検出・CSV出力）
- Rev2: 不可視文字除去＆OB行除去機能追加
- Rev3: PE項目・職番変換ロジック修正
- Rev4: 同乗者(onb)集計処理追加
- Rev5: Excel出力ルーチン追加（基本書式設定）
- Rev6: 全セル上揃え＆折り返し設定
- Rev7: 電話番号セルの折り返し無効化＆日付行背景色設定
- Rev8: ヘッダー行書式調整（左寄せ・上下二重罫線）
- Rev9: 同乗者の所属一致時に*付与、*付き同乗者をExcelでハイライト

"""

import pandas as pd
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side


def run(schedule_input, emp_input, config_path=None):
    """
    Streamlit から呼び出されるエントリポイント。
    schedule_input, emp_input はファイルパス文字列または file-like オブジェクト(Uploader)のどちらでも可。
    """
    # 1) pandas で読み込む
    if hasattr(schedule_input, "read"):
        schedule_df = pd.read_csv(schedule_input, header=None, dtype=str).fillna("")
    else:
        schedule_df = pd.read_csv(schedule_input, header=None, dtype=str).fillna("")

    if hasattr(emp_input, "read"):
        emp_df = pd.read_csv(emp_input, header=None, dtype=str).fillna("")
    else:
        emp_df = pd.read_csv(emp_input, header=None, dtype=str).fillna("")

    # 2) 既存の整形処理をここで実行
    #    （従来 generate_schedule.py で書いていた処理をすべてこの run 内に移動）


# ==== DataFrame読み込み & セルクリーニング ====
def clean_cell(text):
    s = str(text)
    s = re.sub(r"[\u200b\u200c\u200d\u2060\uFEFF\u00A0\t\r\n]", "", s)
    return s.strip()


schedule_df = pd.read_csv(SCHEDULE_CSV, header=None, dtype=str).fillna("")
emp_df = pd.read_csv(EMP_CSV, header=None, dtype=str).fillna("")

# ==== 職員情報整形 ====
emp_df.columns = [f"col_{i+1}" for i in range(emp_df.shape[1])]
emp_df["emp_no"] = emp_df["col_3"].str.zfill(5)
emp_df["full_name"] = emp_df["col_5"].apply(clean_cell) + emp_df["col_7"].apply(
    clean_cell
)
emp_df = emp_df.reset_index().rename(columns={"index": "sort_idx"})


# ==== 前処理関数 ====
def remove_blank_and_ob(df):
    df1 = df[
        ~df.apply(
            lambda r: bool(re.search(r"00099\d{3}", "".join(r.astype(str)))), axis=1
        )
    ]
    for i, row in df1.iterrows():
        if re.match(r"^[A-Z]+[O0]B$", clean_cell(row.iat[0])):
            df1 = df1.iloc[:i]
            break
    df2 = df1[~df1.apply(lambda r: all(clean_cell(c) == "" for c in r), axis=1)]
    return df2.reset_index(drop=True)


# ==== ヘッダー検出 ====
def find_header_rows(df):
    idxs = []
    for i in range(len(df) - 1):
        first = clean_cell(df.iat[i, 0])
        if not re.match(r"^[A-Z]{2,}$", first):
            continue
        nxt = df.iloc[i + 1]
        cnt = sum(
            bool(re.match(r"^(0?[1-9]|[12][0-9]|3[01])$", clean_cell(c))) for c in nxt
        )
        if cnt >= 25:
            idxs.append(i)
    return idxs


schedule_df = remove_blank_and_ob(schedule_df)
headers = find_header_rows(schedule_df)


# ==== ブロック切り出し ====
def slice_blocks(df, hdrs):
    blocks = []
    n = len(df)
    for idx, h in enumerate(hdrs):
        nxt = hdrs[idx + 1] if idx + 1 < len(hdrs) else n
        blocks.append((h, h + 1, nxt))
    return blocks


blocks = slice_blocks(schedule_df, headers)


# ==== ブロック整形 ====
def reshape_block(df, h, d, e):
    raw_hdr = [clean_cell(x) for x in df.iloc[h].tolist()]
    hdr_items = [col for col in raw_hdr if col]
    transformed = []
    for col in hdr_items:
        if "PE有効期限" in col:
            transformed.append("PE")
        elif m := re.match(r"^PE(\d{6})$", col):
            transformed.append(m.group(1))
        elif m := re.match(r"^PE[（\(](\d{6})[）\)]$", col):
            transformed.append(m.group(1))
        elif "社員番号" in col:
            transformed.append("職番")
        else:
            transformed.append(col)
    hdr31 = transformed[:31] + [""] * (31 - len(transformed))
    raw_dates = [clean_cell(x) for x in df.iloc[d].tolist()]
    date_cols = [
        i
        for i, x in enumerate(raw_dates)
        if re.match(r"^(0?[1-9]|[12][0-9]|3[01])$", x)
    ]
    date_cols = date_cols[:31] + [None] * (31 - len(date_cols))
    dr = [
        raw_dates[i] if i is not None and i < len(raw_dates) else "" for i in date_cols
    ]
    rows = df.iloc[d + 1 : e]
    raw_cells = [
        [clean_cell(row[i]) if i is not None else "" for i in date_cols]
        for _, row in rows.iterrows()
    ]
    sched_df = pd.DataFrame(raw_cells)
    return hdr31, date_cols, dr, sched_df, raw_cells


# ==== レコード収集 ====
seen = set()
records = []
for h, d, e in blocks:
    hdr31, date_cols, dr, sched_df, raw_cells = reshape_block(schedule_df, h, d, e)
    emp_no = None
    for cell in hdr31:
        if m := re.search(r"000(\d{5})", cell):
            emp_no = m.group(1)
            break
    if emp_no and emp_no in seen:
        continue
    if emp_no:
        seen.add(emp_no)
    sort_key = float("inf")
    if emp_no:
        mrow = emp_df[emp_df.emp_no == emp_no]
        if not mrow.empty:
            sort_key = int(mrow.iloc[0].sort_idx)
            hdr31[0] = mrow.iloc[0].full_name
            col8 = mrow.iloc[0].col_8
            hdr31[29] = f"PH{col8}" if col8 in ["1", "2", "3"] else col8
            hdr31[30] = mrow.iloc[0].col_1
    hdr31[1] = ""
    merged_sched = [
        "\n".join([v for v in sched_df.iloc[:, i] if v])
        for i in range(sched_df.shape[1])
    ]
    raw_cols = []
    for idx_col in range(len(date_cols)):
        col_vals = [
            row[idx_col] for row in raw_cells if idx_col < len(row) and row[idx_col]
        ]
        raw_cols.append(col_vals)
    records.append(
        {
            "sort": sort_key,
            "hdr": hdr31,
            "dr": dr,
            "sched": merged_sched,
            "raw_cols": raw_cols,
            "onb": [],
        }
    )
records.sort(key=lambda x: x["sort"])

# ==== 同乗者(onb)ロジック（Rev9: 所属一致時に*付与） ====
for rec in records:
    own_aff = rec["hdr"][30]
    merged_onb = []
    for i, code_str in enumerate(rec["sched"]):
        if not re.match(r"^\d", code_str) or re.match(r"^\d+DH$", code_str):
            merged_onb.append("")
            continue
        codes = code_str.split("\n")
        names = []
        for other in records:
            if other is rec:
                continue
            for c in codes:
                if c in other["raw_cols"][i] and other["hdr"][0] not in [
                    n.rstrip("*") for n in names
                ]:
                    name = other["hdr"][0]
                    if other["hdr"][30] == own_aff:
                        name = f"{name}*"
                    names.append(name)
        merged_onb.append("\n".join(names))
    rec["onb"] = merged_onb

# ==== CSV 出力 ====
out = []
for rec in records:
    out.extend([rec["hdr"], rec["dr"], rec["sched"], rec["onb"]])
pd.DataFrame(out).to_csv(
    os.path.join(BASE_DIR, "blocks_output.csv"), header=False, index=False
)
print("blocks_output.csv written")


# ==== Excel 出力ルーチン (Rev10: *付き同乗者をハイライト) ====
def write_to_excel(records, output_path):
    wb = Workbook()
    ws = wb.active
    grey_fill = PatternFill(fill_type="solid", fgColor="DDDDDD")
    highlight_fill = PatternFill(
        fill_type="solid", fgColor="FFFF00"
    )  # 黄色でハイライト
    double_side = Side(border_style="double", color="000000")
    double_border = Border(top=double_side, bottom=double_side)
    # 列幅固定設定 (1〜31列)
    for col in range(1, 32):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 8
    row_idx = 1
    for rec in records:
        # ヘッダー行
        for col_idx, val in enumerate(rec["hdr"], start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=False
            )
            cell.border = double_border
        row_idx += 1
        # 日付行
        for col_idx, val in enumerate(rec["dr"], start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.fill = grey_fill
        row_idx += 1
        # スケジュール行
        for col_idx, val in enumerate(rec["sched"], start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
        row_idx += 1
        # 同乗者行
        for col_idx, val in enumerate(rec["onb"], start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
            if "*" in val:
                cell.fill = highlight_fill
        row_idx += 1
    wb.save(output_path)

    # 3) 出力ファイルを保存
    out_csv = "formatted_schedule.csv"
    out_xlsx = "formatted_schedule.xlsx"
    # schedule_df.to_csv(out_csv, index=False)
    # <Excel 保存ロジック>

    return out_csv, out_xlsx


if __name__ == "__main__":
    # コマンドライン実行時の既存挙動を保持
    # BASE_DIR = os.getcwd()
    # SCHEDULE_CSV = os.path.join(BASE_DIR, "schedule.csv")
    # EMP_CSV      = os.path.join(BASE_DIR, "emp_no.csv")
    run("schedule.csv", "emp_no.csv")

# ==== 固定パス設定 ====
# BASE_DIR = os.getcwd()
# SCHEDULE_CSV = os.path.join(BASE_DIR, "schedule.csv")
# EMP_CSV = os.path.join(BASE_DIR, "emp_no.csv")
