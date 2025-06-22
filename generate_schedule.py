import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side

# ==== ヘルパー関数群 ====


def clean_cell(text):
    s = str(text)
    return re.sub(r"[\u200b\u200c\u200d\u2060\uFEFF\u00A0\t\r\n]", "", s).strip()


def remove_blank_and_ob(df):
    # 空セルおよび "nan"、"OB" 行を削除
    df = df.replace({pd.NA: "", "nan": "", "NaN": ""})
    df = df[(df != "").any(axis=1)]
    df = df[~df.apply(lambda row: any(str(x).strip() == "OB" for x in row), axis=1)]
    return df.reset_index(drop=True)


def find_header_rows(df):
    # ヘッダー行: 1列目末尾2文字が英大文字
    hdrs = []
    for i in range(len(df) - 1):
        val = clean_cell(df.iat[i, 0])
        if re.search(r"[A-Z]{2}$", val):
            hdrs.append(i)
    return hdrs


def slice_blocks(df):
    # ヘッダー行検出 & 各ブロックの開始・終了と日付行検出
    hdrs = find_header_rows(df)
    blocks = []
    n = len(df)
    for idx, h in enumerate(hdrs):
        start = h
        end = hdrs[idx + 1] if idx + 1 < len(hdrs) else n
        # 日付行を探す
        for d in range(h + 1, end):
            row = [clean_cell(x) for x in df.iloc[d]]
            # 数字セルのみのリスト
            date_positions = [
                i
                for i, v in enumerate(row[:31])
                if re.fullmatch(r"(0?[1-9]|[12][0-9]|3[01])", v)
            ]
            if len(date_positions) >= 1:
                blocks.append((start, d, end, date_positions))
                break
    return blocks


def reshape_block(df, block, emp_df):
    h, d, e, date_pos = block
    # ヘッダー
    raw_hdr = [clean_cell(x) for x in df.iloc[h]]
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
    hdr31 = hdr31[:31] + [""] * (31 - len(hdr31))

    # 日付行
    row_dates = [clean_cell(x) for x in df.iloc[d]]
    dr = [row_dates[i] for i in date_pos] + [""] * (31 - len(date_pos))

    # スケジュール
    sched = []
    for r in range(d + 1, e):
        row = [clean_cell(x) for x in df.iloc[r]]
        sched.append([row[i] for i in date_pos])

    # 同乗者
    emp_no = re.search(r"(\d{5})$", hdr31[2])
    emp_no = emp_no.group(1) if emp_no else None
    onb = []
    for r in range(d + 1, e):
        extras = [clean_cell(x) for x in df.iloc[r, 31:]]
        names = []
        for name in extras:
            if not name:
                continue
            if emp_no is not None:
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
    br = Border(top=Side("double"), bottom=Side("double"))
    # 列幅
    for c in range(1, 32):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 8
    r = 1
    for rec in records:
        for c, v in enumerate(rec["hdr"], start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(wrap_text=True)
            cell.border = br
        r += 1
        for c, v in enumerate(rec["dr"], start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.fill = grey
        r += 1
        for row_vals in rec["sched"]:
            for c, v in enumerate(row_vals, start=1):
                ws.cell(row=r, column=c, value=v)
            r += 1
        for c, v in enumerate(rec["onb"], start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(wrap_text=True)
            if "*" in v:
                cell.fill = hi
        r += 1
    wb.save(out_xlsx)


def run(schedule_input, emp_input, config_path=None):
    df = (
        pd.read_csv(schedule_input, header=None, dtype=str).fillna("")
        if not hasattr(schedule_input, "read")
        else pd.read_csv(schedule_input, header=None, dtype=str).fillna("")
    )
    emp_df = (
        pd.read_csv(emp_input, header=None, dtype=str).fillna("")
        if not hasattr(emp_input, "read")
        else pd.read_csv(emp_input, header=None, dtype=str).fillna("")
    )
    blocks = slice_blocks(df.copy().map(clean_cell).pipe(remove_blank_and_ob))
    records = [reshape_block(df, b, emp_df) for b in blocks]
    out_csv = "formatted_schedule.csv"
    rows = []
    for r in records:
        rows.extend([r["hdr"], r["dr"]] + r["sched"] + [r["onb"]])
    pd.DataFrame(rows).to_csv(out_csv, index=False, header=False)
    out_xlsx = "formatted_schedule.xlsx"
    write_to_excel(records, out_xlsx)
    return out_csv, out_xlsx


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser()
    p.add_argument("--schedule", default="schedule.csv")
    p.add_argument("--emp", default="emp_no.csv")
    a = p.parse_args()
    run(a.schedule, a.emp)
