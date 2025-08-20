import io
import os
import glob
import inspect
import importlib
import importlib.util
import tempfile
from typing import Any, Dict, Union, Callable

import streamlit as st

# 実行ディレクトリをこのファイルの場所に固定（ローカル挙動に合わせる）
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)


# ============================
# generate_schedule.run の取得
# ============================
def load_run_callable() -> tuple[str, Callable[..., Any]]:
    env_mod = os.environ.get("NAGU_GEN_MODULE")
    candidates = [env_mod] if env_mod else []
    candidates += [
        "generate_schedule",
        "generate_schedule35e",
        "generate_schedule35d",
        "generate_schedule35",
        "generate_schedule_fixed",
        "generate_schedule30c",
        "generate_schedule25c",
    ]
    # 1) 既知モジュール名を順に import
    for name in [c for c in candidates if c]:
        try:
            mod = importlib.import_module(name)
            run_fn = getattr(mod, "run", None)
            if callable(run_fn):
                return name, run_fn
        except Exception:
            pass
    # 2) カレント配下から generate_schedule*.py を探索
    pattern = os.path.join(SCRIPT_DIR, "generate_schedule*.py")
    paths = sorted(glob.glob(pattern), reverse=True)
    for path in paths:
        try:
            modname = os.path.splitext(os.path.basename(path))[0]
            spec = importlib.util.spec_from_file_location(modname, path)
            if spec and spec.loader:
                mod = importlib.util.module_from_spec(spec)
                import sys

                sys.modules[modname] = mod
                spec.loader.exec_module(mod)
                run_fn = getattr(mod, "run", None)
                if callable(run_fn):
                    return modname, run_fn
        except Exception:
            continue
    raise ImportError("run() を含む生成モジュールを読み込めませんでした。")


MODULE_NAME, RUN = load_run_callable()


# ============================
# ヘルパー
# ============================
def _guess_kwargs_for_run(
    sched_path: str,
    pref_path: str,
    sim_path: Union[str, None],
    prev_xlsx_path: Union[str, None],
    compare_flag: bool,
) -> Dict[str, Any]:
    """run() のシグネチャに合わせて引数マップを組み立てる"""
    sig = inspect.signature(RUN)
    param_names = list(sig.parameters.keys())
    kw: Dict[str, Any] = {}

    mapping = {
        "schedule_file": sched_path,
        "schedule": sched_path,
        "sched_file": sched_path,
        "schedule_path": sched_path,
        "pref_file": pref_path,
        "pref": pref_path,
        "pref_path": pref_path,
        "sim_file": sim_path,
        "sim": sim_path,
        "simslot_file": sim_path,
        "simslot_path": sim_path,
        "prev_xlsx": prev_xlsx_path,
        "prev_excel": prev_xlsx_path,
        "prev_file": prev_xlsx_path,
        "previous_xlsx": prev_xlsx_path,
        "compare_prev": bool(compare_flag),
        "diff_mode": bool(compare_flag),
        "compare": bool(compare_flag),
        "is_compare": bool(compare_flag),
    }
    for k in param_names:
        if k in mapping and mapping[k] is not None:
            kw[k] = mapping[k]

    # 3引数のみ等、特殊ケースは素直に位置対応
    if not kw and len(param_names) == 3:
        return {
            param_names[0]: sched_path,
            param_names[1]: pref_path,
            param_names[2]: sim_path,
        }
    return kw


def _save_temp_from_upload(uploaded, suffix: str) -> str:
    """アップロードファイルを変換せずバイナリのまま一時保存し、パスを返す"""
    if uploaded is None:
        return ""
    raw = uploaded.getvalue()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=SCRIPT_DIR)
    tmp.write(raw)
    tmp.flush()
    tmp.close()
    return tmp.name


def _to_bytes(path_or_bytes: Union[str, bytes, bytearray, io.BytesIO, None]) -> bytes:
    """run() の戻り値（パス/バイト/バッファ）を bytes に正規化"""
    if path_or_bytes is None:
        return b""
    if isinstance(path_or_bytes, (bytes, bytearray)):
        return bytes(path_or_bytes)
    if isinstance(path_or_bytes, io.BytesIO):
        path_or_bytes.seek(0)
        return path_or_bytes.read()
    if isinstance(path_or_bytes, str):
        with open(path_or_bytes, "rb") as f:
            return f.read()
    if hasattr(path_or_bytes, "read"):
        return path_or_bytes.read()  # type: ignore
    return bytes(path_or_bytes)


# ============================
# UI
# ============================
st.set_page_config(
    page_title="✈ NAGU 乗務割整形支援ツール（STAGE）",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🧪 STAGE：NAGU 乗務割整形支援ツール")

with st.sidebar:
    st.markdown("### 入力ファイル")
    sched_up = st.file_uploader(
        "スケジュールCSV（必須）", type=["csv"], key="sched_csv"
    )
    pref_up = st.file_uploader("PREF.xlsx（必須）", type=["xlsx"], key="pref_xlsx")
    sim_up = st.file_uploader("SIM Slot List（任意）", type=["xlsx"], key="sim_xlsx")

    st.divider()
    st.markdown(
        """
**ご注意（文字化け対策）**
- スケジュールCSVは **UTF-8(BOM)** または **CP932(Shift-JIS)** のいずれかに統一してください。
- Excel で「CSV UTF-8 (コンマ区切り)(*.csv)」として保存するのが確実ですが「開かない」のが一番です。。
- 異なる文字コードが同一ファイルに混在していると文字化けの原因になります。
-ダウンロードしたcsvファイルはスプレッドシートなどで開かずそのまま使うか、そのまま転送してください。
        """
    )

    st.divider()
    compare = st.toggle("前回Excelと比較する（任意）", value=False)
    prev_up = None
    if compare:
        prev_up = st.file_uploader("前回Excel（.xlsx）", type=["xlsx"], key="prev_xlsx")

    st.divider()
    run_btn = st.button("🚀 実行", use_container_width=True)

# Main
if run_btn:
    if not sched_up or not pref_up:
        st.error("スケジュールCSV と PREF.xlsx は必須です。")
        st.stop()

    try:
        # すべて実ファイルパスで run() に渡す（ローカル挙動と統一）
        sched_path = _save_temp_from_upload(sched_up, ".csv")
        pref_path = _save_temp_from_upload(pref_up, ".xlsx")
        sim_path = _save_temp_from_upload(sim_up, ".xlsx") if sim_up else None
        prev_path = (
            _save_temp_from_upload(prev_up, ".xlsx") if (compare and prev_up) else None
        )

        kwargs = _guess_kwargs_for_run(
            sched_path, pref_path, sim_path, prev_path, compare_flag=compare
        )
        result = RUN(**kwargs)

        # 出力の正規化（bytes化）
        csv_bytes: bytes = b""
        xlsx_bytes: bytes = b""
        csv_name: str = "schedule.csv"
        xlsx_name: str = "schedule.xlsx"

        if isinstance(result, (tuple, list)) and len(result) == 2:
            csv_src, xlsx_src = result
            if isinstance(csv_src, str):
                csv_name = os.path.basename(csv_src)
            if isinstance(xlsx_src, str):
                xlsx_name = os.path.basename(xlsx_src)
            csv_bytes = _to_bytes(csv_src)
            xlsx_bytes = _to_bytes(xlsx_src)
        elif isinstance(result, dict):
            csv_src = (
                result.get("csv") or result.get("csv_bytes") or result.get("csv_path")
            )
            xlsx_src = (
                result.get("xlsx")
                or result.get("xlsx_bytes")
                or result.get("xlsx_path")
            )
            if isinstance(result.get("csv_name"), str):
                csv_name = result["csv_name"]
            if isinstance(result.get("xlsx_name"), str):
                xlsx_name = result["xlsx_name"]
            csv_bytes = _to_bytes(csv_src) if csv_src is not None else b""
            xlsx_bytes = _to_bytes(xlsx_src) if xlsx_src is not None else b""
        else:
            raise RuntimeError(
                "run() は (csv, xlsx) のタプルまたは dict を返す必要があります。"
            )

        st.success("処理が完了しました。ダウンロードしてください。")

        # Excel（xlsx）は必ず生バイトで配布（破損防止）
        if xlsx_bytes:
            st.download_button(
                label="📥 Excel（.xlsx）をダウンロード",
                data=xlsx_bytes,
                file_name=xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("Excel 出力が見つかりませんでした。")

        # CSV は補助的に配布（UTF-8 BOM / CP932 の整備は run 側の仕様に依存）
        if csv_bytes:
            st.download_button(
                label="CSV（そのまま）",
                data=csv_bytes,
                file_name=csv_name,
                mime="text/csv",
            )
        else:
            st.info("CSV 出力が見つかりませんでした。")

    except Exception as e:
        st.error(f"処理中にエラーが発生しました：{e}")
