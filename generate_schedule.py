import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side

# ==== ヘルパー関数群 ====


def clean_cell(text):
    """セル内の不要な不可視文字を削除してトリム"""
    s = str(text)
    return re.sub(r"[\u200b\u200c\u200d\u2060\uFEFF\u00A0\t\r\n]", "", s).strip()


def remove_blank_and_ob(df):
    """
    空行（全セル空）および"OB"のみの行を削除
    """
    df = df.replace({pd.NA: "", "nan": "", "NaN": ""})
    # 全セルが空文字の行を除去
    df = df[(df != "").any(axis=1)]
    # OB だけの行を除去
    df = df[~df.apply(lambda row: all(str(c).strip() == "OB" for c in row), axis=1)]
    return df.reset_index(drop=True)


def find_header_rows(df):
    """
    ヘッダー行を検出: 1列目の末尾2文字が英大文字
    """
    hdrs = []
    for idx, row in df.iterrows():
        val = clean_cell(row.iloc[0])
        if re.search(r"[A-Z]{2}$", val):
            hdrs.append(idx)
    return hdrs


def find_date_row(df, start_idx):
    """
    ヘッダー行 start_idx の次行以降で、先頭31列に日付(1-31)が並ぶ行を探す
    見つかればインデックス、なければ None
    """
    for i in range(start_idx + 1, len(df)):
        row = [clean_cell(x) for x in df.iloc[i].tolist()][:31]
        # 数字または空
        if all(
            (re.fullmatch(r"(0?[1-9]|[12][0-9]|3[01])", cell) or cell == "")
            for cell in row
        ) and any(re.fullmatch(r"(0?[1-9]|[12][0-9]|3[01])", cell) for cell in row):
            return i
    return None


def slice_blocks(df):
    """
    ヘッダー行からブロック毎の (h, date_row, next_h) タプルを作成
    """
    hdrs = find_header_rows(df)
    blocks = []
    total = len(df)
    for idx, h in enumerate(hdrs):
        nxt = hdrs[idx + 1] if idx + 1 < len(hdrs) else total
        d = find_date_row(df, h)
        if d is not None:
            blocks.append((h, d, nxt))
    return blocks


def reshape_block(df, h, d, e, emp_df):
    """
    ブロックの範囲 h(ヘッダー), d(日付行), e(次ヘッダーor終端) で整形
    """
    # ヘッダー
    raw_hdr = [clean_cell(x) for x in df.iloc[h].tolist()][:31]
    hdr31 = []
    for col in raw_hdr:
        if "PE有効期限" in col:
            hdr31.append("PE")
        elif m := re.match(r"^PE[（\(]?(\d{6})[）\)]?", col):
            hdr31.append(m.group(1))
        elif "社員番号" in col:
            hdr31.append("職番")
        else:
            hdr31.append(col)
    hdr31 += [""] * (31 - len(hdr31))

    # 日付行
    row_dates = [clean_cell(x) for x in df.iloc[d].tolist()][:31]
    dr = row_dates

    # スケジュール行
    sched = []
    for r in range(d + 1, e):
        row_vals = [clean_cell(x) for x in df.iloc[r].tolist()][:31]
        sched.append(row_vals)

    # 同乗者行
    emp_no = None
    if len(hdr31) > 2:
        m = re.search(r"(\d{5})$", hdr31[2])
        emp_no = m.group(1) if m else None
    onb = []
    for r in range(d + 1, e):
        extras = [clean_cell(x) for x in df.iloc[r].tolist()[31:]]
        names = []
        for name in extras:
            if not name:
                continue
            if emp_no:
                recs = emp_df[emp_df.iloc[:, 2] == emp_no]
                if not recs.empty and recs.iloc[0, 3] == recs.iloc[0, 3]:
                    name += "*"
            names.append(name)
        onb.append("\n".join(names))

    return {"hdr": hdr31, "dr": dr, "sched": sched, "onb": onb}


def write_to_excel(records, out_xlsx):
    wb = Workbook()
    ws = wb.active
    grey = PatternFill(fill_type="solid", fgColor="DDDDDD")
    hi = PatternFill(fill_type="solid", fgColor="FFFF00")
    dbl = Side(border_style="double", color="000000")
    brd = Border(top=dbl, bottom=dbl)

    # 列幅
    for c in range(1, 32):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 8

    row_idx = 1
    for rec in records:
        for c, val in enumerate(rec["hdr"], start=1):
            cell = ws.cell(row=row_idx, column=c, value=val)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
            cell.border = brd
        row_idx += 1
        for c, val in enumerate(rec["dr"], start=1):
            cell = ws.cell(row=row_idx, column=c, value=val)
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.fill = grey
        row_idx += 1
        for row_vals in rec["sched"]:
            for c, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=row_idx, column=c, value=val)
                cell.alignment = Alignment(
                    horizontal="left", vertical="top", wrap_text=True
                )
            row_idx += 1
        for c, val in enumerate(rec["onb"], start=1):
            cell = ws.cell(row=row_idx, column=c, value=val)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
            if "*" in val:
                cell.fill = hi
        row_idx += 1
    wb.save(out_xlsx)


# ==== エントリポイント ====
def run(schedule_input, emp_input, config_path=None):
    # 入力読み込み
    sched_df = (
        pd.read_csv(schedule_input, header=None, dtype=str).fillna("")
        if not hasattr(schedule_input, "read")
        else pd.read_csv(schedule_input, header=None, dtype=str).fillna("")
    )
    emp_df = (
        pd.read_csv(emp_input, header=None, dtype=str).fillna("")
        if not hasattr(emp_input, "read")
        else pd.read_csv(emp_input, header=None, dtype=str).fillna("")
    )

    # ブロック検出・整形
    df = sched_df.copy().map(clean_cell).pipe(remove_blank_and_ob)
    blocks = slice_blocks(df)
    records = [reshape_block(df, h, d, e, emp_df) for h, d, e in blocks]

    # CSV出力
    out_csv = "formatted_schedule.csv"
    rows = []
    for rec in records:
        rows.extend([rec["hdr"], rec["dr"]] + rec["sched"] + [rec["onb"]])
    pd.DataFrame(rows).to_csv(out_csv, index=False, header=False)

    # Excel出力
    out_xlsx = "formatted_schedule.xlsx"
    write_to_excel(records, out_xlsx)

    return out_csv, out_xlsx


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--schedule", default="schedule.csv")
    parser.add_argument("--emp", default="emp_no.csv")
    args = parser.parse_args()
    run(args.schedule, args.emp)
