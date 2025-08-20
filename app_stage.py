import io
import os
import glob
import inspect
import importlib
import importlib.util
import tempfile
import traceback
from typing import Any, Dict, Union, Callable

import streamlit as st

# Ensure working directory = this script's folder (to mimic LOCAL behavior)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)


# ============================
# Dynamic loader for `run()`
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
    for name in [c for c in candidates if c]:
        try:
            mod = importlib.import_module(name)
            if hasattr(mod, "run") and callable(getattr(mod, "run")):
                return name, getattr(mod, "run")
        except Exception:
            pass

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
                if hasattr(mod, "run") and callable(getattr(mod, "run")):
                    return modname, getattr(mod, "run")
        except Exception:
            continue
    raise ImportError("run() を含む生成モジュールを読み込めませんでした。")


MODULE_NAME, RUN = load_run_callable()


# ============================
# Helpers
# ============================
def _guess_kwargs_for_run(
    sched_path: str,
    pref_path: str,
    sim_path: Union[str, None],
    prev_xlsx_path: Union[str, None],
    compare_flag: bool,
) -> Dict[str, Any]:
    sig = inspect.signature(RUN)
    param_names = list(sig.parameters.keys())
    kw: Dict[str, Any] = {}

    # Map by common names; always pass PATHS (not BytesIO) to mimic LOCAL
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

    # If nothing matched and exactly 3 params, pass positionally (sched, pref, sim)
    if not kw and len(param_names) == 3:
        return {
            param_names[0]: sched_path,
            param_names[1]: pref_path,
            param_names[2]: sim_path,
        }
    return kw


def _save_temp_binary(data: bytes, suffix: str, *, dir_: str = SCRIPT_DIR) -> str:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=dir_)
    tmp.write(data)
    tmp.flush()
    tmp.close()
    return tmp.name


def _save_temp_from_upload(uploaded, suffix: str) -> str:
    """バイナリのまま保存（変換なし）"""
    raw = uploaded.getvalue()
    return _save_temp_binary(raw, suffix, dir_=SCRIPT_DIR)


def _to_bytes(path_or_bytes: Union[str, bytes, bytearray, io.BytesIO, None]) -> bytes:
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


def _stat(path: Union[str, None]) -> dict:
    if not path:
        return {"exists": False}
    try:
        stt = os.stat(path)
        return {"exists": True, "size": stt.st_size, "path": path}
    except FileNotFoundError:
        return {"exists": False, "path": path}
    except Exception as e:
        return {"exists": False, "path": path, "error": str(e)}


# ---------- 混在エンコーディング対応（UTF-8＋SJIS） ----------
def _is_sjis_pair(b1: int, b2: int) -> bool:
    # SJIS: lead 0x81–0x9F or 0xE0–0xFC, trail 0x40–0x7E or 0x80–0xFC（0x7F除く）
    return ((0x81 <= b1 <= 0x9F) or (0xE0 <= b1 <= 0xFC)) and (
        (0x40 <= b2 <= 0x7E) or (0x80 <= b2 <= 0xFC)
    )


def _mixed_utf8_sjis_to_text(raw: bytes) -> str:
    """
    先頭から走査し、UTF-8 で切れる所は UTF-8 として読む。
    そこで読めなければ SJIS の2バイト並びを優先的に解釈。
    どちらでもなければ置換（'�'）で前進。
    """
    out = []
    i = 0
    n = len(raw)
    while i < n:
        # UTF-8 を 1〜4バイトで試す
        got = False
        for l in (1, 2, 3, 4):
            if i + l <= n:
                try:
                    ch = raw[i : i + l].decode("utf-8")
                    out.append(ch)
                    i += l
                    got = True
                    break
                except UnicodeDecodeError:
                    pass
        if got:
            continue
        # SJIS の2バイト並びを試す
        if i + 1 < n and _is_sjis_pair(raw[i], raw[i + 1]):
            try:
                out.append(bytes([raw[i], raw[i + 1]]).decode("cp932"))
            except Exception:
                out.append("�")
            i += 2
            continue
        # ASCII 単バイトはそのまま、その他は置換
        b = raw[i]
        out.append(chr(b) if b < 0x80 else "�")
        i += 1
    return "".join(out)


def _ensure_cp932_csv_path_from_upload(uploaded) -> tuple[str, str]:
    """
    アップロードCSVをミックス復号→最終的に CP932 で再エンコードして
    一時CSVに保存し、そのパスを返す。
    戻り値: (csv_path, note)
    """
    raw = uploaded.getvalue()
    text = _mixed_utf8_sjis_to_text(raw)
    # 改行をLFに正規化（必要なければ外してOK）
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    try:
        data = text.encode("cp932", errors="strict")
    except Exception:
        data = text.encode("cp932", errors="replace")
    tmp_path = _save_temp_binary(data, ".csv", dir_=SCRIPT_DIR)
    return tmp_path, "normalized: mixed UTF-8/SJIS → CP932"


# ============================
# UI
# ============================
st.set_page_config(
    page_title="✈ NAGU 乗務割整形支援ツール（STAGE, Mixed-encoding Normalizer）",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🧪 STAGE（混在エンコーディング補正 → CP932 正規化）")

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

st.markdown("#### 実行環境")
st.code(
    f"CWD: {os.getcwd()}\nSCRIPT_DIR: {SCRIPT_DIR}\nMODULE: {MODULE_NAME}\nRUN signature: {inspect.signature(RUN)}",
    language="bash",
)

# Main
if run_btn:
    if not sched_up or not pref_up:
        st.error("スケジュールCSV と PREF.xlsx は必須です。")
        st.stop()

    # ★ CSV は「混在エンコーディング補正 → CP932 正規化」してから run() に渡す
    sched_path, sched_note = _ensure_cp932_csv_path_from_upload(sched_up)

    # 他のファイルはバイナリのまま保存
    pref_path = _save_temp_from_upload(pref_up, ".xlsx")
    sim_path = _save_temp_from_upload(sim_up, ".xlsx") if sim_up else None
    prev_path = (
        _save_temp_from_upload(prev_up, ".xlsx") if (compare and prev_up) else None
    )

    # Show saved paths and existence
    st.markdown("#### 入力ファイルの保存状態")
    st.json(
        {
            "schedule": _stat(sched_path) | {"note": sched_note},
            "pref": _stat(pref_path),
            "sim": _stat(sim_path),
            "prev_xlsx": _stat(prev_path),
        }
    )

    try:
        kwargs = _guess_kwargs_for_run(
            sched_path, pref_path, sim_path, prev_path, compare_flag=compare
        )
        st.markdown("#### run() に渡す引数")
        st.json({k: (v if isinstance(v, bool) else str(v)) for k, v in kwargs.items()})

        result = RUN(**kwargs)

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

        st.success("処理が完了しました。ダウンロードしてください。")
        if xlsx_bytes:
            st.download_button(
                label="📥 Excel（.xlsx）をダウンロード",
                data=xlsx_bytes,
                file_name=xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        if csv_bytes:
            st.download_button(
                label="CSV（生のまま）",
                data=csv_bytes,
                file_name=csv_name,
                mime="text/csv",
            )

    except FileNotFoundError as fnf:
        st.error(f"FileNotFoundError: {fnf}")
        st.code("\\n".join(os.listdir(os.getcwd())))
        st.code(traceback.format_exc())
    except Exception as e:
        st.error(f"処理中にエラーが発生しました：{e}")
        st.code("\\n".join(os.listdir(os.getcwd())))
        st.code(traceback.format_exc())
