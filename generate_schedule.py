import pandas as pd
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side

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


# ==== DataFrame読み込み & セルクリーニング ====
def clean_cell(text):
    s = str(text)
    s = re.sub(r"[\u200b\u200c\u200d\u2060\uFEFF\u00A0\t\r\n]", "", s)
    return s.strip()


def remove_blank_and_ob(df):
    # OB(欠員)行・空行を削除
    mask = []
    for _, row in df.iterrows():
        row_str = "".join(row.tolist())
        mask.append(bool(re.search(r"[01]\-\d{2}\.\d{2}|OB", row_str)))
    return df[mask].reset_index(drop=True)


def find_header_rows(df):
    # 1列目が2文字以上大文字英字の行をヘッダー行とみなす
    hdrs = []
    for idx, row in df.iterrows():
        if re.match(r"^[A-Z]{2,}$", str(row.iloc[0])):
            hdrs.append(idx)
    return hdrs


def slice_blocks(df, hdrs):
    # 各ヘッダー行から次のヘッダー行までをひとブロックとする
    blocks = []
    n = len(df)
    for i, h in enumerate(hdrs):
        nxt = hdrs[i + 1] if i + 1 < len(hdrs) else n
        blocks.append((h, h + 1, nxt))
    return blocks


def reshape_block(df, h, d, e):
    """
    df: 全体DataFrame
    h: ヘッダー行 idx, d: 日付行 idx, e: 次ヘッダー行 idx（終端）
    emp_df: モジュールグローバルから参照
    """
    # ヘッダー行のクレンジング＆変換
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

    # 日付行
    dr = [clean_cell(x) for x in df.iloc[d].tolist()][:31]

    # スケジュール行：h+2 から e-1 まで
    sched = []
    for ri in range(h + 2, e):
        sched.append([clean_cell(x) for x in df.iloc[ri].tolist()][:31])

    # 同乗者行 onb: 列31以降の同乗者名をまとめて改行連結
    merged_onb = []
    for ri in range(h + 2, e):
        names = []
        for name in [clean_cell(x) for x in df.iloc[ri, 31:].tolist()]:
            if not name:
                continue
            # 同所属に * を付与
            own_aff = hdr31[2]  # 職番列
            emp_no = re.sub(r".*?(\d{5})$", r"\1", hdr31[2])
            # モジュールの emp_df グローバルを使う
            recs = emp_df[emp_df.iloc[:, 2] == emp_no]
            if not recs.empty and recs.iloc[0, 3] == recs.iloc[0, 3]:  # 所属照合
                name = name + "*"
            names.append(name)
        merged_onb.append("\n".join(names))

    return {"hdr": hdr31, "dr": dr, "sched": sched, "onb": merged_onb}


# ==== Excel 出力ルーチン（Rev9対応） ====
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
        # ヘッダー行
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
    """
    :param schedule_input: ファイルパス(str) または UploadedFile
    :param emp_input:      ファイルパス(str) または UploadedFile
    :param config_path:    未使用（将来設定対応用）
    :return: (out_csv_path, out_xlsx_path)
    """
    # --- 1) 読み込み ---
    if hasattr(schedule_input, "read"):
        schedule_df_local = pd.read_csv(schedule_input, header=None, dtype=str).fillna(
            ""
        )
    else:
        schedule_df_local = pd.read_csv(schedule_input, header=None, dtype=str).fillna(
            ""
        )
    if hasattr(emp_input, "read"):
        emp_df_local = pd.read_csv(emp_input, header=None, dtype=str).fillna("")
    else:
        emp_df_local = pd.read_csv(emp_input, header=None, dtype=str).fillna("")

    # モジュール関数内で emp_df を参照できるようにグローバルに設定
    global emp_df
    emp_df = emp_df_local

    # --- 2) ブロック生成＆整形 ---
    df = schedule_df_local.copy().applymap(clean_cell).pipe(remove_blank_and_ob)
    headers = find_header_rows(df)
    blocks = slice_blocks(df, headers)
    records = [reshape_block(df, h, d, e) for h, d, e in blocks]

    # --- 3) CSV 出力 ---
    out_csv = "blocks_output.csv"
    out_list = []
    for rec in records:
        out_list.extend([rec["hdr"], rec["dr"], rec["sched"], rec["onb"]])
    pd.DataFrame(out_list).to_csv(out_csv, header=False, index=False)

    # --- 4) Excel 出力 ---
    out_xlsx = "blocks_output.xlsx"
    write_to_excel(records, out_xlsx)

    return out_csv, out_xlsx


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="スケジュール整形ツール CLI")
    parser.add_argument(
        "--schedule", default="schedule.csv", help="入力スケジュールCSVのパス"
    )
    parser.add_argument("--emp", default="emp_no.csv", help="入力職員番号CSVのパス")
    args = parser.parse_args()
    run(args.schedule, args.emp)
