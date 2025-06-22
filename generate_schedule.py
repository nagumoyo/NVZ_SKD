import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side

# ==== ヘルパー関数群 ====


def clean_cell(text):
    """セル内の不要な不可視文字を削除し、両端をトリム"""
    s = str(text)
    return re.sub(r"[\u200b\u200c\u200d\u2060\uFEFF\u00A0\t\r\n]", "", s).strip()


def remove_blank_and_ob(df):
    """
    空セル（全セルが空文字）の行およびOBのみの行を削除する
    """
    df = df.replace({pd.NA: "", "nan": "", "NaN": ""})
    # 全セルが空文字の行を除去
    df = df[(df != "").any(axis=1)]
    # OBのみの行を除去
    df = df[~df.apply(lambda row: all(str(c).strip() == "OB" for c in row), axis=1)]
    return df.reset_index(drop=True)


def find_header_rows(df):
    """
    ヘッダー行を検出: 1列目の値の末尾2文字が英大文字かつ次行に日付候補が含まれる
    """
    hdrs = []
    for i in range(len(df) - 1):
        first = clean_cell(df.iat[i, 0])
        if not re.search(r"[A-Z]{2}$", first):
            continue
        next_row = [clean_cell(x) for x in df.iloc[i + 1].tolist()]
        cnt = sum(bool(re.fullmatch(r"(0?[1-9]|[12][0-9]|3[01])", v)) for v in next_row)
        if cnt >= 1:
            hdrs.append(i)
    return hdrs


def slice_blocks(df):
    """
    ヘッダー行リストから各ブロック(h, d, next_h, date_positions)を作成
    """
    hdrs = find_header_rows(df)
    blocks = []
    total = len(df)
    for idx, h in enumerate(hdrs):
        nxt = hdrs[idx + 1] if idx + 1 < len(hdrs) else total
        for i in range(h + 1, nxt):
            row = [clean_cell(x) for x in df.iloc[i].tolist()]
            date_positions = [
                j
                for j, v in enumerate(row)
                if re.fullmatch(r"(0?[1-9]|[12][0-9]|3[01])", v)
            ]
            if date_positions:
                date_positions = date_positions[:31]
                blocks.append((h, i, nxt, date_positions))
                break
    return blocks


def reshape_block(df, h, d, e, date_positions, emp_df):
    """
    ブロックごとにヘッダー、日付行、スケジュール行、同乗者行を抽出してdictで返す
    keys: hdr, dr, sched, onb
    """
    # ヘッダー
    raw_hdr = [clean_cell(x) for x in df.iloc[h].tolist()]
    hdr31 = raw_hdr[:31] + [""] * (31 - len(raw_hdr[:31]))
    # 日付行
    dr = [clean_cell(df.iat[d, j]) for j in date_positions] + [""] * (
        31 - len(date_positions)
    )
    # スケジュール行
    sched = [
        [clean_cell(df.iat[r, j]) for j in date_positions] for r in range(d + 1, e)
    ]
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
    # 列幅設定
    for c in range(1, 32):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 8
    r = 1
    for rec in records:
        for c, v in enumerate(rec["hdr"], start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
            cell.border = brd
        r += 1
        for c, v in enumerate(rec["dr"], start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.fill = grey
        r += 1
        for row_vals in rec["sched"]:
            for c, v in enumerate(row_vals, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.alignment = Alignment(
                    horizontal="left", vertical="top", wrap_text=True
                )
            r += 1
        for c, v in enumerate(rec["onb"], start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
            if "*" in v:
                cell.fill = hi
        r += 1
    wb.save(out_xlsx)


# ==== エントリポイント ====
def run(schedule_input, emp_input, config_path=None):
    if hasattr(schedule_input, "read"):
        sched_df = pd.read_csv(schedule_input, header=None, dtype=str).fillna("")
    else:
        sched_df = pd.read_csv(schedule_input, header=None, dtype=str).fillna("")
    if hasattr(emp_input, "read"):
        emp_df = pd.read_csv(emp_input, header=None, dtype=str).fillna("")
    else:
        emp_df = pd.read_csv(emp_input, header=None, dtype=str).fillna("")
    df = sched_df.copy().map(clean_cell).pipe(remove_blank_and_ob)
    blocks = slice_blocks(df)
    records = [
        reshape_block(df, h, d, e, date_pos, emp_df) for h, d, e, date_pos in blocks
    ]
    out_csv = "formatted_schedule.csv"
    rows = []
    for rec in records:
        rows.extend([rec["hdr"], rec["dr"]] + rec["sched"] + [rec["onb"]])
    pd.DataFrame(rows).to_csv(out_csv, index=False, header=False)
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
