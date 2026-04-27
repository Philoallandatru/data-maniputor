from __future__ import annotations

import base64
import hashlib
import sys
from io import BytesIO
from numbers import Number
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

APP_TITLE = "PandasAI + 本地 LLM Excel 分析助手"
DEFAULT_API_BASE = "http://localhost:1234/v1"
DEFAULT_MODEL_NAME = "qwen-local"
DEFAULT_API_KEY = "lm-studio"
DEFAULT_TEMPERATURE = 0.1
DEFAULT_MAX_TOKENS = 4096
DEFAULT_MAX_ROWS = 50_000
PREVIEW_ROWS = 100
CHARTS_DIR = Path("exports") / "charts"
SUPPORTED_FILE_TYPES = ("csv", "xlsx", "xlsm", "xls")
CSV_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "latin1")
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}


def init_session_state() -> None:
    st.session_state.setdefault("messages", [])
    st.session_state.setdefault("active_dataset_key", None)


def normalize_api_base_url(api_base: str) -> str:
    normalized = (api_base or "").strip().rstrip("/")
    if not normalized:
        raise ValueError("API Base URL 不能为空。")
    return normalized


def normalize_lmstudio_model_name(model_name: str) -> str:
    normalized = (model_name or "").strip()
    if not normalized:
        raise ValueError("模型名不能为空。")
    if not normalized.startswith("openai/"):
        normalized = f"openai/{normalized}"
    return normalized


def ensure_chart_dir() -> None:
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)


def import_pandasai():
    try:
        import pandasai as pai
    except ImportError as exc:
        raise RuntimeError(
            "未检测到 pandasai。请先安装 requirements.txt 中的依赖。"
        ) from exc

    try:
        from pandasai_litellm.litellm import LiteLLM
    except Exception as exc:
        raise RuntimeError(
            "未能导入 pandasai-litellm。请确认使用 Python 3.11，且已安装 "
            "`pandasai` 与 `pandasai-litellm` 的兼容版本。"
        ) from exc

    return pai, LiteLLM


def configure_pandasai_llm(
    api_base: str,
    model_name: str,
    api_key: str,
    temperature: float,
    max_tokens: int,
    save_charts: bool,
):
    pai, LiteLLM = import_pandasai()

    normalized_api_base = normalize_api_base_url(api_base)
    normalized_model = normalize_lmstudio_model_name(model_name)
    normalized_api_key = (api_key or "").strip()

    if not normalized_api_key:
        raise ValueError("API Key 不能为空，可填写任意非空字符串。")

    ensure_chart_dir()

    llm = LiteLLM(
        model=normalized_model,
        api_base=normalized_api_base,
        api_key=normalized_api_key,
        temperature=float(temperature),
        max_tokens=int(max_tokens),
    )

    config_candidates = [
        {
            "llm": llm,
            "verbose": True,
            "save_logs": True,
            "max_retries": 2,
            "save_charts": bool(save_charts),
            "save_charts_path": str(CHARTS_DIR),
            "enable_cache": False,
        },
        {
            "llm": llm,
            "verbose": True,
            "save_logs": True,
            "max_retries": 2,
        },
        {"llm": llm},
    ]

    applied_config = None
    config_manager = getattr(pai, "config", None)
    if config_manager is not None and hasattr(config_manager, "set"):
        last_error = None
        for candidate in config_candidates:
            try:
                config_manager.set(candidate)
                applied_config = candidate
                break
            except Exception as exc:
                last_error = exc
        if applied_config is None and last_error is not None:
            raise RuntimeError(f"PandasAI 配置失败：{last_error}") from last_error
    else:
        applied_config = config_candidates[0]

    setattr(pai, "_codex_pandasai_config", applied_config)
    setattr(pai, "_codex_save_charts", bool(save_charts))
    setattr(pai, "_codex_api_base", normalized_api_base)
    setattr(pai, "_codex_model_name", normalized_model)
    return pai


def to_pandasai_df(pai, df: pd.DataFrame):
    errors: list[str] = []

    dataframe_cls = getattr(pai, "DataFrame", None)
    if dataframe_cls is not None:
        try:
            return dataframe_cls(df.copy(), _table_name="uploaded_data")
        except Exception as exc:
            errors.append(f"pai.DataFrame(df) 失败: {exc}")

    smart_dataframe_cls = getattr(pai, "SmartDataframe", None)
    if smart_dataframe_cls is not None:
        config = getattr(pai, "_codex_pandasai_config", None)
        try:
            if config:
                return smart_dataframe_cls(df.copy(), name="uploaded_data", config=config)
            return smart_dataframe_cls(df.copy(), name="uploaded_data")
        except Exception as exc:
            errors.append(f"SmartDataframe(df) 失败: {exc}")

    details = "\n".join(f"- {item}" for item in errors) if errors else "- 无可用 DataFrame 包装器"
    raise RuntimeError(f"PandasAI DataFrame 包装失败：\n{details}")


def _is_excel_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in {".xlsx", ".xlsm", ".xls"}


def _excel_engine(filename: str) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"
    raise ValueError(f"不支持的 Excel 扩展名：{suffix}")


@st.cache_data(show_spinner=False)
def list_excel_sheets(file_bytes: bytes, filename: str) -> list[str]:
    with BytesIO(file_bytes) as buffer:
        excel_file = pd.ExcelFile(buffer, engine=_excel_engine(filename))
        return [str(name) for name in excel_file.sheet_names]


def _read_csv_with_fallback(file_bytes: bytes, nrows: int | None) -> pd.DataFrame:
    attempts: list[str] = []

    for encoding in CSV_ENCODINGS:
        try:
            return pd.read_csv(
                BytesIO(file_bytes),
                nrows=nrows,
                encoding=encoding,
                low_memory=False,
            )
        except Exception as exc:
            attempts.append(f"encoding={encoding}: {exc}")

    for encoding in CSV_ENCODINGS:
        try:
            return pd.read_csv(
                BytesIO(file_bytes),
                nrows=nrows,
                encoding=encoding,
                sep=None,
                engine="python",
            )
        except Exception as exc:
            attempts.append(f"encoding={encoding}, sep=auto: {exc}")

    attempt_text = "\n".join(f"- {item}" for item in attempts[-8:])
    raise ValueError(f"CSV 读取失败，请检查编码或分隔符。\n{attempt_text}")


@st.cache_data(show_spinner=False)
def read_uploaded_file(
    file_bytes: bytes,
    filename: str,
    sheet_name: str | None,
    max_rows: int,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    suffix = Path(filename).suffix.lower()
    if suffix not in {".csv", ".xlsx", ".xlsm", ".xls"}:
        raise ValueError("文件格式不支持，请上传 CSV / XLSX / XLSM / XLS 文件。")

    nrows = max_rows + 1 if max_rows and max_rows > 0 else None
    selected_sheet = sheet_name or "CSV"

    if suffix == ".csv":
        df = _read_csv_with_fallback(file_bytes, nrows=nrows)
    else:
        with BytesIO(file_bytes) as buffer:
            excel_file = pd.ExcelFile(buffer, engine=_excel_engine(filename))
            resolved_sheet = sheet_name or str(excel_file.sheet_names[0])
            df = pd.read_excel(excel_file, sheet_name=resolved_sheet, nrows=nrows)
            selected_sheet = str(resolved_sheet)

    truncated = bool(max_rows and len(df) > max_rows)
    if truncated:
        df = df.head(max_rows).copy()

    df = df.copy()
    df.columns = [str(column) for column in df.columns]

    return df, {"sheet_name": selected_sheet, "truncated": truncated}


def build_column_profile(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "column": [str(column) for column in df.columns],
            "dtype": [str(dtype) for dtype in df.dtypes],
            "non_null": df.notna().sum().tolist(),
            "null": df.isna().sum().tolist(),
        }
    )


def build_default_prompt(df: pd.DataFrame, sheet_name: str) -> str:
    preview_columns = "、".join(str(column) for column in df.columns[:30])
    if len(df.columns) > 30:
        preview_columns += " ..."

    return f"""你是 SSD 性能测试数据分析助手。请基于当前 DataFrame 分析数据。

当前 sheet：{sheet_name}
当前行数：{len(df)}
当前字段：{preview_columns}

请先识别可能的版本列、平台列、workload 列、指标列和数值列。

如果数据中包含多个固件版本，请对比不同版本之间的性能变化。

重点关注：
1. 带宽 / IOPS 是否下降
2. latency / P95 / P99 是否上升
3. 哪些 workload、平台、容量组合退化最明显
4. 给出一个简洁的结论表格

请用中文回答。"""


def _is_local_image_path(value: Any) -> bool:
    if isinstance(value, Path):
        return value.exists() and value.suffix.lower() in IMAGE_SUFFIXES
    if isinstance(value, str):
        path = Path(value)
        return path.exists() and path.suffix.lower() in IMAGE_SUFFIXES
    return False


def _is_data_image_uri(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("data:image/")


def _decode_data_image_uri(value: str) -> bytes:
    _, encoded = value.split(",", 1)
    return base64.b64decode(encoded)


def _prepare_response_for_storage(response_value: Any, save_charts: bool) -> Any:
    if _is_data_image_uri(response_value):
        return _decode_data_image_uri(response_value)

    if _is_local_image_path(response_value) and not save_charts:
        return Path(response_value).read_bytes()

    return response_value


def _summarize_for_debug(value: Any, limit: int = 3000) -> str:
    try:
        if isinstance(value, pd.DataFrame):
            preview = value.head(20).to_markdown(index=False)
        elif isinstance(value, pd.Series):
            preview = value.head(20).to_string()
        else:
            preview = repr(value)
    except Exception:
        preview = f"<无法预览对象: {type(value).__name__}>"

    return preview if len(preview) <= limit else f"{preview[:limit]} ...[truncated]"


def render_response(response: Any) -> None:
    if response is None:
        st.info("模型没有返回可展示的结果。")
        return

    if isinstance(response, bytes):
        st.image(response, use_container_width=True)
        return

    if isinstance(response, pd.Series):
        st.dataframe(response.to_frame(name=response.name or "value"), use_container_width=True)
        return

    if isinstance(response, pd.DataFrame):
        st.dataframe(response, use_container_width=True)
        return

    if isinstance(response, Number) and not isinstance(response, bool):
        st.metric("结果", f"{response}")
        return

    if _is_local_image_path(response):
        st.image(str(response), use_container_width=True)
        st.caption(f"图表路径：`{response}`")
        return

    if _is_data_image_uri(response):
        st.image(_decode_data_image_uri(response), use_container_width=True)
        return

    if isinstance(response, str):
        st.markdown(response)
        return

    if isinstance(response, (list, dict)):
        try:
            st.json(response, expanded=False)
        except Exception:
            st.write(response)
        return

    st.write(response)


def _render_debug_info(debug_info: dict[str, Any]) -> None:
    with st.expander("原始结果 / 调试信息", expanded=False):
        st.write(f"响应类型：`{debug_info.get('response_class', 'unknown')}`")
        if debug_info.get("api_base"):
            st.write(f"API Base：`{debug_info['api_base']}`")
        if debug_info.get("model_name"):
            st.write(f"模型名：`{debug_info['model_name']}`")
        if debug_info.get("last_code_executed"):
            st.code(debug_info["last_code_executed"], language="python")
        st.code(debug_info.get("raw_preview", "<empty>"), language="text")


def _analysis_error_message(exc: Exception) -> str:
    detail = str(exc).strip() or exc.__class__.__name__
    return f"""分析失败，请检查：
1. 本地模型服务是否已启动
2. API Base 是否包含 `/v1`
3. 模型名是否和 `/v1/models` 返回一致
4. 是否已安装 `pandasai`、`pandasai-litellm`
5. 当前 Python 版本是否为 3.11
6. 当前模型是否具备代码生成能力

原始错误：
```text
{detail}
```"""


def _file_error_message(exc: Exception) -> str:
    detail = str(exc).strip() or exc.__class__.__name__

    if "xlrd" in detail.lower():
        return (
            "读取 `.xls` 文件失败，请确认已经安装 `xlrd`。\n\n"
            f"原始错误：\n```text\n{detail}\n```"
        )

    return f"文件读取失败，请检查文件格式、sheet 名称或编码。\n\n```text\n{detail}\n```"


def _render_chat_history() -> None:
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            if message["role"] == "user":
                st.markdown(message["content"])
                continue

            if message.get("is_error"):
                st.error(message["content"])
            else:
                render_response(message["content"])

            debug_info = message.get("debug")
            if debug_info:
                _render_debug_info(debug_info)


def _run_analysis(
    df: pd.DataFrame,
    prompt: str,
    api_base: str,
    model_name: str,
    api_key: str,
    temperature: float,
    max_tokens: int,
    save_charts: bool,
) -> tuple[Any, dict[str, Any]]:
    pai = configure_pandasai_llm(
        api_base=api_base,
        model_name=model_name,
        api_key=api_key,
        temperature=temperature,
        max_tokens=max_tokens,
        save_charts=save_charts,
    )
    pai_df = to_pandasai_df(pai, df)
    raw_response = pai_df.chat(prompt)

    response_value = getattr(raw_response, "value", raw_response)
    prepared_value = _prepare_response_for_storage(response_value, save_charts=save_charts)
    last_code_executed = getattr(raw_response, "last_code_executed", None) or getattr(
        pai_df, "last_code_executed", None
    )

    debug_info = {
        "response_class": type(raw_response).__name__,
        "api_base": getattr(pai, "_codex_api_base", None),
        "model_name": getattr(pai, "_codex_model_name", None),
        "last_code_executed": last_code_executed,
        "raw_preview": _summarize_for_debug(response_value),
    }
    return prepared_value, debug_info


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_session_state()

    st.title(APP_TITLE)
    st.caption("本地运行，不调用 OpenAI 云端接口。适用于 Excel / CSV 的 SSD 性能测试分析与通用数据问答。")

    if sys.version_info >= (3, 12):
        st.warning(
            "当前检测到 Python 3.12+。`pandasai-litellm` 当前通常需要 Python 3.11，"
            "如遇到安装或导入失败，请切换到 Python 3.11 环境。"
        )

    with st.sidebar:
        st.header("本地模型配置")
        api_base = st.text_input("API Base URL", value=DEFAULT_API_BASE)
        model_name = st.text_input("模型名", value=DEFAULT_MODEL_NAME)
        api_key = st.text_input(
            "API Key",
            value=DEFAULT_API_KEY,
            type="password",
            help="可填写任意非空字符串，例如 lm-studio / ollama / EMPTY。",
        )
        temperature = st.number_input(
            "temperature",
            min_value=0.0,
            max_value=2.0,
            value=DEFAULT_TEMPERATURE,
            step=0.1,
        )
        max_tokens = st.number_input(
            "max_tokens",
            min_value=256,
            max_value=32768,
            value=DEFAULT_MAX_TOKENS,
            step=256,
        )
        max_rows = st.number_input(
            "最大读取行数",
            min_value=100,
            max_value=500000,
            value=DEFAULT_MAX_ROWS,
            step=1000,
        )
        save_charts = st.checkbox("是否保存图表", value=True)

        try:
            normalized_api_base = normalize_api_base_url(api_base)
            normalized_model = normalize_lmstudio_model_name(model_name)
            st.caption(f"实际请求 API Base：`{normalized_api_base}`")
            st.caption(f"实际请求模型名：`{normalized_model}`")
        except ValueError:
            st.caption("请先填写合法的 API Base URL 和模型名。")

        st.info("模型名可通过本地服务的 `/v1/models` 查看。")

        if st.button("清空聊天历史", use_container_width=True):
            st.session_state.messages = []
            st.rerun()

    uploaded_file = st.file_uploader(
        "上传 CSV / Excel 文件",
        type=list(SUPPORTED_FILE_TYPES),
        help="支持 .csv / .xlsx / .xlsm / .xls，默认最多读取前 50,000 行。",
    )

    if uploaded_file is None:
        st.info("请先上传一个 CSV 或 Excel 文件。")
        return

    file_bytes = uploaded_file.getvalue()
    if not file_bytes:
        st.error("上传文件为空，请重新选择文件。")
        return

    file_size_mb = len(file_bytes) / (1024 * 1024)
    if file_size_mb > 100:
        st.warning("当前文件超过 100 MB，建议先裁剪后再分析，以降低页面卡顿和模型执行时间。")

    selected_sheet = None
    if _is_excel_file(uploaded_file.name):
        try:
            sheet_names = list_excel_sheets(file_bytes, uploaded_file.name)
        except Exception as exc:
            st.error(_file_error_message(exc))
            return

        if not sheet_names:
            st.error("Excel 文件中未读取到任何 sheet。")
            return

        selected_sheet = st.selectbox("选择要分析的 sheet", sheet_names, index=0)

    file_fingerprint = hashlib.md5(file_bytes).hexdigest()[:12]
    dataset_key = (
        f"{uploaded_file.name}|{file_fingerprint}|{selected_sheet or 'CSV'}|{int(max_rows)}"
    )
    if st.session_state.active_dataset_key != dataset_key:
        st.session_state.active_dataset_key = dataset_key
        st.session_state.messages = []

    try:
        df, metadata = read_uploaded_file(
            file_bytes=file_bytes,
            filename=uploaded_file.name,
            sheet_name=selected_sheet,
            max_rows=int(max_rows),
        )
    except Exception as exc:
        st.error(_file_error_message(exc))
        return

    if df.empty:
        st.error("DataFrame 为空，请确认文件内容、sheet 或分隔符是否正确。")
        return

    if metadata["truncated"]:
        st.warning(f"文件超过最大读取行数，当前仅加载前 {int(max_rows):,} 行参与分析。")

    sheet_label = metadata["sheet_name"]
    profile_df = build_column_profile(df)

    metric_cols = st.columns(4)
    metric_cols[0].metric("行数", f"{len(df):,}")
    metric_cols[1].metric("列数", f"{len(df.columns):,}")
    metric_cols[2].metric("当前 sheet", sheet_label)
    metric_cols[3].metric("文件大小", f"{file_size_mb:.1f} MB")

    with st.expander("字段信息", expanded=True):
        st.dataframe(profile_df, use_container_width=True)

    with st.expander("数据预览（前 100 行）", expanded=True):
        st.dataframe(df.head(PREVIEW_ROWS), use_container_width=True)

    default_prompt = build_default_prompt(df, sheet_label)
    with st.expander("默认分析 Prompt", expanded=False):
        st.code(default_prompt, language="text")

    action_cols = st.columns([1, 3])
    run_default_prompt = action_cols[0].button("运行默认分析", use_container_width=True)
    action_cols[1].caption(
        "示例问题：对比 FW880 和 FW881，哪些 workload 性能下降超过 5%？"
    )

    st.subheader("聊天分析")
    if not st.session_state.messages:
        st.info("上传数据后即可提问。支持文本、表格和图表结果。")

    _render_chat_history()

    chat_prompt = st.chat_input(
        "请输入分析问题，例如：对比 FW880 和 FW881，哪些 workload 性能下降超过 5%？",
        disabled=df is None,
    )
    prompt = default_prompt if run_default_prompt else chat_prompt

    if not prompt:
        return

    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("正在调用 PandasAI 和本地模型分析数据..."):
            try:
                response, debug_info = _run_analysis(
                    df=df,
                    prompt=prompt,
                    api_base=api_base,
                    model_name=model_name,
                    api_key=api_key,
                    temperature=float(temperature),
                    max_tokens=int(max_tokens),
                    save_charts=save_charts,
                )
            except Exception as exc:
                error_message = _analysis_error_message(exc)
                st.error(error_message)
                st.session_state.messages.append(
                    {
                        "role": "assistant",
                        "content": error_message,
                        "is_error": True,
                    }
                )
                return

        render_response(response)
        _render_debug_info(debug_info)
        st.session_state.messages.append(
            {
                "role": "assistant",
                "content": response,
                "debug": debug_info,
                "is_error": False,
            }
        )


if __name__ == "__main__":
    main()
