import streamlit as st
import pandas as pd
import re
import io

st.title("整形済スケジュール v15 Webアプリ版")

st.markdown("### ファイルをアップロードしてください")

schedule_file = st.file_uploader("スケジュールCSVファイル", type="csv")
emp_file = st.file_uploader("職員情報ファイル (emp_no.csv)", type="csv")
output_filename = st.text_input("出力ファイル名（例：整形スケジュール_v15.xlsx）", "整形スケジュール_v15.xlsx")

if st.button("出力"):
    if schedule_file is None or emp_file is None:
        st.error("スケジュールファイルと職員情報ファイルの両方をアップロードしてください。")
    else:
        # 読込
        emp_df = pd.read_csv(emp_file, header=None, encoding="utf-8")
        emp_df.columns = [f"col_{i+1}" for i in range(emp_df.shape[1])]
        emp_df["emp_no"] = emp_df["col_3"].astype(str).str.zfill(5)
        emp_df["full_name"] = emp_df["col_5"].astype(str) + emp_df["col_7"].astype(str)

        schedule_df = pd.read_csv(schedule_file, header=None, encoding="utf-8", dtype=str)
        schedule_df.fillna("", inplace=True)

        def is_OB_row(row):
            first_cell = row[0].strip()
            if re.match(r"^[A-Z]+OB$", first_cell):
                return True
            for cell in row:
                if re.search(r"(00099\d{3})", str(cell)):
                    return True
            return False

        header_indices_valid = []
        for i in range(schedule_df.shape[0] - 1):
            first_cell = schedule_df.iloc[i, 0].strip()
            next_row = schedule_df.iloc[i + 1, :].tolist()
            header_match = bool(re.match(r"^[A-Z]{2,}$|^[A-Z]{3,}$", first_cell))
            date_count = sum(1 for cell in next_row if re.match(r"^(0?[1-9]|[12][0-9]|3[01])$", str(cell).strip()))
            if header_match and date_count >= 25:
                header_indices_valid.append(i)

        crew_blocks = []
        for idx, header_idx in enumerate(header_indices_valid):
            next_header_idx = header_indices_valid[idx + 1] if idx + 1 < len(header_indices_valid) else schedule_df.shape[0]
            if idx + 1 < len(header_indices_valid):
                block_end = next_header_idx
            else:
                block_end = schedule_df.shape[0]
                for j in range(header_idx + 2, schedule_df.shape[0]):
                    row = schedule_df.iloc[j, :].tolist()
                    first_cell = row[0].strip()
                    next_row = schedule_df.iloc[j + 1, :].tolist() if j + 1 < schedule_df.shape[0] else []
                    is_fake_header = bool(re.match(r"^[A-Z]{2,}$|^[A-Z]{3,}$", first_cell)) and \
                                     (sum(1 for cell in next_row if re.match(r"^(0?[1-9]|[12][0-9]|3[01])$", str(cell).strip())) < 25)
                    is_blank_row = all([str(cell).strip() == "" for cell in row])
                    if is_OB_row(row) or is_fake_header or is_blank_row:
                        block_end = max(j - 1, header_idx + 2)
                        break
                else:
                    block_end = schedule_df.shape[0]
            header_row = schedule_df.iloc[header_idx, :].tolist()
            if is_OB_row(header_row):
                continue
            crew_blocks.append((header_idx, header_idx + 1, block_end))

        output_rows = []
        for block in crew_blocks:
            header_raw = schedule_df.iloc[block[0], :].tolist()
            emp_no_match = next((re.search(r"(000\d{5})", cell).group(1) for cell in header_raw if re.search(r"(000\d{5})", cell)), None)
            emp_name = ""
            if emp_no_match:
                emp_name_row = emp_df.loc[emp_df["emp_no"] == emp_no_match[-5:], "full_name"]
                if not emp_name_row.empty:
                    emp_name = emp_name_row.iloc[0]
            name_first_cell = header_raw[0].strip()
            final_name = emp_name if emp_name else name_first_cell
            two_letter = header_raw[2].strip() if len(header_raw) > 2 else ""
            header_elements = [final_name, two_letter, emp_no_match if emp_no_match else ""]
            header_elements += [cell for i, cell in enumerate(header_raw) if i not in [0,2] and str(cell).strip() != ""]

            date_row = schedule_df.iloc[block[1], :].tolist()

            sched_search_range = schedule_df.iloc[block[1]+1 : block[2]+1, :]
            block_start_dynamic_idx = None
            for offset, (_, row) in enumerate(sched_search_range.iterrows()):
                if not all([str(cell).strip() == "" for cell in row.tolist()]):
                    block_start_dynamic_idx = block[1] + 1 + offset
                    break
            if block_start_dynamic_idx is not None:
                sched_rows = schedule_df.iloc[block_start_dynamic_idx : block[2]+1, :]
            else:
                sched_rows = pd.DataFrame()

            date_col_indices = [i for i, cell in enumerate(date_row) if re.match(r"^(0?[1-9]|[12][0-9]|3[01])$", str(cell).strip())]
            date_col_indices = date_col_indices[:31]

            header_row_fixed = [header_elements[i] if i < len(header_elements) else "" for i in range(31)]
            date_row_fixed = [date_row[i] for i in date_col_indices]
            while len(date_row_fixed) < 31:
                date_row_fixed.append("")

            merged_schedule_row = [""] * 31
            for i, col_idx in enumerate(date_col_indices):
                sched_texts = sched_rows.iloc[:, col_idx].apply(lambda x: str(x)).tolist() if not sched_rows.empty else []
                merged_schedule_row[i] = "\n".join(sched_texts)

            output_rows.append(header_row_fixed)
            output_rows.append(date_row_fixed)
            output_rows.append(merged_schedule_row)

        output_df = pd.DataFrame(output_rows)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, header=False)
        buffer.seek(0)

        st.success("出力が完了しました！ 下のボタンからダウンロードしてください。")
        st.download_button(
            label="📥 ダウンロード",
            data=buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
