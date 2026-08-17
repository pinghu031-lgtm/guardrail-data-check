# -*- coding: utf-8 -*-
"""Streamlit Community Cloud 入口。"""

import hmac
import os

import streamlit as st

from cloud_service import (
    UploadValidationError,
    UploadedFileData,
    process_uploads,
    upload_fingerprint,
    validate_upload_metadata,
)
from guardrail_core import APP_VERSION


st.set_page_config(
    page_title=f"护栏数据异常检查工具 {APP_VERSION}",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 3rem;}
    [data-testid="stMetric"] {background: #f7f9fc; border: 1px solid #d9e2f3; border-radius: 8px; padding: 12px;}
    .app-note {color: #5b6573; font-size: 0.92rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def configured_password() -> str:
    password = os.environ.get("APP_PASSWORD", "")
    if password:
        return password
    try:
        return str(st.secrets.get("APP_PASSWORD", ""))
    except Exception:
        return ""


def require_access() -> None:
    expected = configured_password()
    if not expected:
        return
    if st.session_state.get("authenticated"):
        return

    st.title("护栏数据异常检查工具")
    st.caption(f"云端共享版 · {APP_VERSION}")
    supplied = st.text_input("访问口令", type="password", placeholder="请输入访问口令")
    if st.button("进入工具", type="primary", width="stretch"):
        if hmac.compare_digest(supplied, expected):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("访问口令错误。")
    st.stop()


require_access()

st.title("护栏数据异常检查工具")
st.caption(f"云端共享版 · {APP_VERSION}")
st.markdown(
    '<p class="app-note">上传 Excel 后在云端临时处理；处理结果通过浏览器下载，文件不会写入 GitHub 仓库。</p>',
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "选择 Excel 文件",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="单次最多 20 个文件；单文件不超过 50 MiB；总大小不超过 200 MiB。",
)

upload_validation_error = ""
uploaded_payloads = []
if uploaded_files:
    try:
        validate_upload_metadata([int(item.size) for item in uploaded_files])
    except UploadValidationError as exc:
        upload_validation_error = str(exc)
    else:
        uploaded_payloads = [
            UploadedFileData(item.name, item.getvalue()) for item in uploaded_files
        ]

current_upload_signature = upload_fingerprint(uploaded_payloads) if uploaded_payloads else None
if st.session_state.get("processed_upload_signature") != current_upload_signature:
    st.session_state.pop("process_result", None)

if uploaded_files:
    st.write(f"已选择 **{len(uploaded_files)}** 个文件")
    st.dataframe(
        [
            {"文件名": item.name, "大小（MiB）": round(item.size / 1024 / 1024, 2)}
            for item in uploaded_files
        ],
        hide_index=True,
        width="stretch",
    )
if upload_validation_error:
    st.error(upload_validation_error)

if st.button(
    "开始处理",
    type="primary",
    width="stretch",
    disabled=not uploaded_payloads,
):
    st.session_state.pop("process_result", None)
    st.session_state.pop("processed_upload_signature", None)
    try:
        with st.spinner("正在检查并生成修改数据，请稍候……"):
            st.session_state["process_result"] = process_uploads(uploaded_payloads)
            st.session_state["processed_upload_signature"] = current_upload_signature
        st.success("处理完成。")
    except UploadValidationError as exc:
        st.session_state.pop("process_result", None)
        st.session_state.pop("processed_upload_signature", None)
        st.error(str(exc))
    except Exception:
        st.session_state.pop("process_result", None)
        st.session_state.pop("processed_upload_signature", None)
        st.error("处理失败。请确认文件未损坏且表头符合要求；如问题持续存在，请联系工具维护者。")

result = st.session_state.get("process_result")
if result is not None:
    st.subheader("处理统计")
    metric_columns = st.columns(4)
    metric_columns[0].metric("成功文件", result.stats.get("files_success", 0))
    metric_columns[1].metric("失败文件", result.stats.get("files_failed", 0))
    metric_columns[2].metric("异常区间", len(result.preview_rows))
    metric_columns[3].metric("修改单元格", result.modification_stats.get("modified_cells", 0))

    st.subheader("结果预览")
    if result.preview_rows:
        st.dataframe(result.preview_rows, hide_index=True, width="stretch")
    else:
        st.info("未发现符合当前规则的异常。")

    download_columns = st.columns(2)
    download_columns[0].download_button(
        "下载异常检查结果",
        data=result.result_excel,
        file_name=result.result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    download_columns[1].download_button(
        "下载已修改数据 ZIP",
        data=result.modified_zip,
        file_name=result.modified_filename,
        mime="application/zip",
        width="stretch",
    )

    with st.expander("运行日志"):
        st.code("\n".join(result.logs) or "无日志", language="text")

st.divider()
st.caption("支持 .xlsx 和 .xlsm。请勿上传与护栏数据检查无关或包含不应外传信息的文件。")
