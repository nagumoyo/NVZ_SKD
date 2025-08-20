import io
import os
import glob
import inspect
import importlib
import importlib.util
import tempfile
from typing import Any, Dict, Union, Callable

import streamlit as st


# ============================
# Dynamic loader for `run()`
# ============================
def load_run_callable() -> Callable[..., Any]:
    """
    Try to import a function named `run` from the user's generator module.
    """
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

    for name in [c for c in candidates if c]:
        try:
            mod = importlib.import_module(name)
            if hasattr(mod, "run") and callable(getattr(mod, "run")):
                return getattr(mod, "run")
        except Exception:
            pass

    here = os.path.dirname(os.path.abspath(__file__))
    pattern = os.path.join(here, "generate_schedule*.py")
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
                if hasattr(mod, "run") and callable(getattr(mod, "run")):
                    return getattr(mod, "run")
        except Exception:
            continue

    raise ImportError(
        "run() を含む生成モジュールを読み込めませんでした。"
        "NAGU_GEN_MODULE 環境変数にモジュール名（例: generate_schedule35e）を設定するか、"
        "app_stage.py と同じフォルダに generate_schedule*.py を配置してください。"
    )


RUN = load_run_callable()


# ============================
# Helpers
# ============================
def _guess_kwargs_for_run(
    sched: Union[io.BytesIO, str],
    pref: Union[io.BytesIO, str],
    sim: Union[io.BytesIO, str, None],
    prev_xlsx: Union[io.BytesIO, str, None],
    compare_flag: bool,
) -> Dict[str, Any]:
    sig = inspect.signature(RUN)
    param_names = list(sig.parameters.keys())
    kw: Dict[str, Any] = {}

    sched_keys = ["schedule_file", "schedule", "sched_file", "schedule_path"]
    pref_keys = ["pref_file", "pref", "pref_path"]
    sim_keys = ["sim_file", "sim", "simslot_file", "simslot_path"]
    prev_keys = ["prev_xlsx", "prev_excel", "prev_file", "previous_xlsx"]
    compare_keys = ["compare_prev", "diff_mode", "compare", "is_compare"]

    for k in sched_keys:
        if k in param_names:
            kw[k] = sched
            break
    for k in pref_keys:
        if k in param_names:
            kw[k] = pref
            break
    for k in sim_keys:
        if k in param_names and sim is not None:
            kw[k] = sim
            break
    for k in prev_keys:
        if k in param_names and prev_xlsx is not None:
            kw[k] = prev_xlsx
            break
    for k in compare_keys:
        if k in param_names:
            kw[k] = bool(compare_flag)
            break

    if not kw and len(param_names) == 3:
        return {param_names[0]: sched, param_names[1]: pref, param_names[2]: sim}
    return kw


def _to_bytes(path_or_bytes: Union[str, bytes, bytearray, io.BytesIO, None]) -> bytes:
    if path_or_bytes is None:
        return b""
    if isinstance(path_or_bytes, bytes):
        return path_or_bytes
    if isinstance(path_or_bytes, bytearray):
        return bytes(path_or_bytes)
    if isinstance(path_or_bytes, io.BytesIO):
        path_or_bytes.seek(0)
        return path_or_bytes.read()
    if isinstance(path_or_bytes, str):
        with open(path_or_bytes, "rb") as f:
            return f.read()
    if hasattr(path_or_bytes, "read"):
        return path_or_bytes.read()  # type: ignore[attr-defined]
    return bytes(path_or_bytes)


def _read_csv_text(csv_source: Union[str, bytes, bytearray, io.BytesIO]) -> str:
    raw: bytes
    if isinstance(csv_source, (bytes, bytearray)):
        raw = bytes(csv_source)
    elif isinstance(csv_source, io.BytesIO):
        csv_source.seek(0)
        raw = csv_source.read()
    elif isinstance(csv_source, str):
        with open(csv_source, "rb") as f:
            raw = f.read()
    else:
        try:
            raw = csv_source.read()  # type: ignore[attr-defined]
        except Exception:
            raw = b""

    for enc in ("utf-8-sig", "utf-8", "cp932", "utf-16", "utf-16le", "utf-16be"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", errors="replace")


def _save_temp(data: Union[io.BytesIO, bytes, bytearray], suffix: str) -> str:
    if isinstance(data, io.BytesIO):
        data.seek(0)
        content = data.read()
    elif isinstance(data, (bytes, bytearray)):
        content = bytes(data)
    else:
        raise TypeError("Unsupported type for _save_temp")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(content)
    tmp.flush()
    tmp.close()
    return tmp.name


def _normalize_csv_upload(
    uploaded: "st.runtime.uploaded_file_manager.UploadedFile",
) -> io.BytesIO:
    """
    Normalize user-uploaded CSV bytes to UTF-8 (no BOM) so that pandas default decoding works.
    """
    raw = uploaded.getvalue()
    # Try common JP encodings; fall back to 'replace'
    for enc in ("utf-8-sig", "utf-8", "cp932", "utf-16", "utf-16le", "utf-16be"):
        try:
            text = raw.decode(enc)
            break
        except Exception:
            text = None
    if text is None:
        text = raw.decode("utf-8", errors="replace")
    # Normalize newlines
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Re-encode to UTF-8 (no BOM). Pandas default will decode it correctly.
    return io.BytesIO(text.encode("utf-8"))


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
        # Normalize CSV to UTF-8 for robust downstream decoding
        sched_io = _normalize_csv_upload(sched_up)

        # Excel files: just pass through
        pref_io = io.BytesIO(pref_up.getvalue())
        sim_io = io.BytesIO(sim_up.getvalue()) if sim_up else None
        prev_io = io.BytesIO(prev_up.getvalue()) if (compare and prev_up) else None

        # First attempt: call run with BytesIOs (preferred for Streamlit)
        kwargs = _guess_kwargs_for_run(
            sched_io, pref_io, sim_io, prev_io, compare_flag=compare
        )
        try:
            result = RUN(**kwargs)
        except TypeError:
            # Fallback to temp files
            sched_path = _save_temp(sched_io, ".csv")
            pref_path = _save_temp(pref_io, ".xlsx")
            sim_path = _save_temp(sim_io, ".xlsx") if sim_io else None
            prev_path = _save_temp(prev_io, ".xlsx") if prev_io else None

            kwargs2 = _guess_kwargs_for_run(
                sched_path, pref_path, sim_path, prev_path, compare_flag=compare
            )
            result = RUN(**kwargs2)

        # Normalize result
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
                "run() の戻り値形式が想定外です。tuple(list)かdictを返してください。"
            )

        # ============================
        # Download buttons
        # ============================
        st.success("処理が完了しました。ダウンロードしてください。")

        if xlsx_bytes:
            st.download_button(
                label="📥 Excel（.xlsx）をダウンロード",
                data=xlsx_bytes,
                file_name=xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("Excel 出力が見つかりませんでした。")

        if csv_bytes:
            csv_text = _read_csv_text(csv_bytes)
            csv_utf8_sig = csv_text.encode("utf-8-sig")
            try:
                csv_cp932 = csv_text.encode("cp932", errors="strict")
            except Exception:
                csv_cp932 = csv_text.encode("cp932", errors="replace")

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="CSV（UTF-8 BOM付き）",
                    data=csv_utf8_sig,
                    file_name=os.path.splitext(csv_name)[0] + "_utf8.csv",
                    mime="text/csv",
                )
            with col2:
                st.download_button(
                    label="CSV（Excel向けCP932）",
                    data=csv_cp932,
                    file_name=os.path.splitext(csv_name)[0] + "_cp932.csv",
                    mime="text/csv",
                )
        else:
            st.info("CSV 出力が見つかりませんでした。")

    except Exception as e:
        st.error(f"処理中にエラーが発生しました：{e}")
