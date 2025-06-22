import pandas as pd
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side


# ==== ヘルパー関数群 ====
def clean_cell(text):
    s = str(text)
    s = re.sub(r"[\u200b\u200c\u200d\u2060\uFEFF\u00A0\t\r\n]", "", s)
    return s.strip()


def remove_blank_and_ob(df):
    """
    空行（すべてのセルが空文字列）および OB を含む行を削除します。
    """
    # 1) 全セルが空文字列の行を削除
    df = df[(df != "").any(axis=1)]
    # 2) 行内に "OB" を含む行を削除
    df = df[
        ~df.apply(lambda row: any(str(cell).strip() == "OB" for cell in row), axis=1)
    ]
    return df.reset_index(drop=True)


def find_header_rows(df):
    hdrs = []
    for idx, row in df.iterrows():
        if re.match(r"^[A-Z]{2,}$", str(row.iloc[0])):
            hdrs.append(idx)
    return hdrs


def find_date_row(df, start_idx):
    for i in range(start_idx + 1, len(df)):
        vals = [clean_cell(x) for x in df.iloc[i].tolist()]
        date_cells = vals[:31]
        if all((re.fullmatch(r"\d{1,2}", v) or v == "") for v in date_cells) and any(
            re.fullmatch(r"\d{1,2}", v) for v in date_cells
        ):
            return i
    raise IndexError(f"Date row not found after header at row {start_idx}")


def reshape_block(df, h, d, e):
    # ヘッダー行
    raw_hdr = [clean_cell(x) for x in df.iloc[h].tolist()]
    transformed = []
    for col in raw_hdr:
        if "PE有効期限" in col:
            transformed.append("PE")
        elif m := re.match(r"^PE[（\(](\d{6})[）\)]$", col):
            transformed.append(m.group(1))
        elif "社員番号" in col:
            transformed.append("職番")
        else:
            transformed.append(col)
    hdr31 = transformed[:31] + [""] * (31 - len(transformed))

    # 日付行
    dr = [clean_cell(x) for x in df.iloc[d].tolist()][:31]

    # スケジュール行
    sched = []
    for ri in range(d + 1, e):
        sched.append([clean_cell(x) for x in df.iloc[ri].tolist()][:31])

    # 同乗者行
    merged_onb = []
    emp_no = re.sub(r".*?(\d{5})$", r"\1", hdr31[2])
    for ri in range(d + 1, e):
        names = []
        for name in [clean_cell(x) for x in df.iloc[ri, 31:].tolist()]:
            if not name:
                continue
            recs = emp_df_global[emp_df_global.iloc[:, 2] == emp_no]
            if not recs.empty and recs.iloc[0, 3] == recs.iloc[0, 3]:
                name += "*"
            names.append(name)
        merged_onb.append("\n".join(names))

    return {"hdr": hdr31, "dr": dr, "sched": sched, "onb": merged_onb}


def write_to_excel(records, output_path):
    wb = Workbook()
    ws = wb.active
    grey_fill = PatternFill(fill_type="solid", fgColor="DDDDDD")
    highlight_fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    double_side = Side(border_style="double", color="000000")
    double_border = Border(top=double_side, bottom=double_side)

    # 列幅設定
    for col in range(1, 32):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 8

    row_idx = 1
    for rec in records:
        # ヘッダー
        for ci, val in enumerate(rec["hdr"], start=1):
            cell = ws.cell(row_idx, ci, val)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
            cell.border = double_border
        row_idx += 1
        # 日付行
        for ci, val in enumerate(rec["dr"], start=1):
            cell = ws.cell(row_idx, ci, val)
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.fill = grey_fill
        row_idx += 1
        # スケジュール行
        for sched_row in rec["sched"]:
            for ci, val in enumerate(sched_row, start=1):
                cell = ws.cell(row_idx, ci, val)
                cell.alignment = Alignment(
                    horizontal="left", vertical="top", wrap_text=True
                )
            row_idx += 1
        # 同乗者行
        for ci, val in enumerate(rec["onb"], start=1):
            cell = ws.cell(row_idx, ci, val)
            cell.alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )
            if "*" in val:
                cell.fill = highlight_fill
        row_idx += 1

    wb.save(output_path)


# ==== エントリポイント ====
def run(schedule_input, emp_input, config_path=None):
    # 1) 読み込み
    if hasattr(schedule_input, "read"):
        schedule_df = pd.read_csv(schedule_input, header=None, dtype=str).fillna("")
    else:
        schedule_df = pd.read_csv(schedule_input, header=None, dtype=str).fillna("")

    if hasattr(emp_input, "read"):
        emp_df = pd.read_csv(emp_input, header=None, dtype=str).fillna("")
    else:
        emp_df = pd.read_csv(emp_input, header=None, dtype=str).fillna("")

    # グローバル変数設定
    global emp_df_global
    emp_df_global = emp_df

    # 2) 整形処理
    df = schedule_df.copy().applymap(clean_cell).pipe(remove_blank_and_ob)
    headers = find_header_rows(df)
    records = []
    for i, h in enumerate(headers):
        next_h = headers[i + 1] if i + 1 < len(headers) else len(df)
        try:
            d = find_date_row(df, h)
        except IndexError:
            continue
        records.append(reshape_block(df, h, d, next_h))

    # 3) CSV 出力
    out_csv = "formatted_schedule.csv"
    out_list = []
    for rec in records:
        out_list.extend([rec["hdr"], rec["dr"]] + rec["sched"] + [rec["onb"]])
    pd.DataFrame(out_list).to_csv(out_csv, header=False, index=False)

    # 4) Excel 出力
    out_xlsx = "formatted_schedule.xlsx"
    write_to_excel(records, out_xlsx)

    return out_csv, out_xlsx


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="スケジュール整形ツール CLI")
    parser.add_argument(
        "--schedule", default="schedule.csv", help="スケジュールCSVのパス"
    )
    parser.add_argument("--emp", default="emp_no.csv", help="職員番号CSVのパス")
    args = parser.parse_args()
    run(args.schedule, args.emp)
