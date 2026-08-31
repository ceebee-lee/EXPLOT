from __future__ import annotations

import hashlib
import logging
import tempfile
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from PO_FOC_LOT_배정_자동화 import (
    APP_VERSION,
    InputFormatError,
    create_allocation_file,
    inspect_inputs,
)


MAX_UPLOAD_BYTES = 50 * 1024 * 1024
MAX_PREVIEW_ROWS = 20_000
KOREA_TZ = ZoneInfo("Asia/Seoul")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


st.set_page_config(
    page_title="PO·FOC LOT 배정",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .block-container {max-width: 1480px; padding-top: 2.2rem; padding-bottom: 4rem;}
      [data-testid="stMetric"] {
        background: #f7f9fc; border: 1px solid #e1e7ef;
        border-radius: 12px; padding: 14px 16px;
      }
      .rule-card {
        background: #f7f9fc; border: 1px solid #e1e7ef;
        border-radius: 12px; padding: 14px 16px; height: 100%;
      }
      .rule-title {font-weight: 700; color: #17365d; margin-bottom: 4px;}
      .rule-text {font-size: 0.92rem; color: #44546a; line-height: 1.45;}
      .muted {color: #667085; font-size: 0.88rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def _fmt_number(value: float | int) -> str:
    number = float(value)
    return f"{int(number):,}" if number.is_integer() else f"{number:,.2f}"


def _file_bytes(uploaded_file) -> bytes:
    return bytes(uploaded_file.getbuffer())


def _validate_upload(uploaded_file, label: str) -> bytes:
    if uploaded_file is None:
        raise InputFormatError(f"{label}을 업로드해 주세요.")
    if Path(uploaded_file.name).suffix.lower() != ".xlsx":
        raise InputFormatError(f"{label}은 .xlsx 파일만 사용할 수 있습니다.")
    payload = _file_bytes(uploaded_file)
    if not payload:
        raise InputFormatError(f"{label}이 비어 있습니다.")
    if len(payload) > MAX_UPLOAD_BYTES:
        raise InputFormatError(f"{label}은 50MB보다 클 수 없습니다.")
    return payload


def _upload_identity(uploaded_file) -> str:
    if uploaded_file is None:
        return "missing"
    file_id = getattr(uploaded_file, "file_id", None)
    size = getattr(uploaded_file, "size", None)
    if file_id:
        return f"{file_id}|{uploaded_file.name}|{size}"
    payload = _file_bytes(uploaded_file)
    return f"{uploaded_file.name}|{len(payload)}|{hashlib.sha256(payload).hexdigest()}"


def _input_signature(order_upload, inventory_upload) -> str:
    digest = hashlib.sha256()
    digest.update(_upload_identity(order_upload).encode("utf-8"))
    digest.update(_upload_identity(inventory_upload).encode("utf-8"))
    return digest.hexdigest()


def _run_signature(
    input_signature: str,
    warehouse_priority: tuple[str, str],
    allow_split: bool,
    order_priority: tuple[str, str],
) -> str:
    digest = hashlib.sha256(input_signature.encode("ascii"))
    digest.update("|".join(warehouse_priority).encode("ascii"))
    digest.update(str(allow_split).encode("ascii"))
    digest.update("|".join(order_priority).encode("ascii"))
    return digest.hexdigest()


def _rule_card(title: str, text: str) -> None:
    st.markdown(
        f'<div class="rule-card"><div class="rule-title">{title}</div>'
        f'<div class="rule-text">{text}</div></div>',
        unsafe_allow_html=True,
    )


def _run_preflight(order_upload, inventory_upload, signature: str) -> dict:
    order_bytes = _validate_upload(order_upload, "PO/FOC 주문 파일")
    inventory_bytes = _validate_upload(inventory_upload, "글로벌 재고 파일")
    with tempfile.TemporaryDirectory(prefix="lot_preflight_") as temp_directory:
        root = Path(temp_directory)
        order_path = root / "orders.xlsx"
        inventory_path = root / "inventory.xlsx"
        order_path.write_bytes(order_bytes)
        inventory_path.write_bytes(inventory_bytes)
        summary = inspect_inputs(order_path, inventory_path)
    summary.update(
        {
            "signature": signature,
            "order_name": Path(order_upload.name).name,
            "inventory_name": Path(inventory_upload.name).name,
        }
    )
    return summary


def _show_preflight(preflight: dict) -> None:
    if preflight["status"] == "PASS":
        st.success("입력 파일 사전검증을 통과했습니다.")
    else:
        st.warning("배정은 가능하지만 확인할 항목이 있습니다.")
        for warning in preflight["warnings"]:
            st.caption(f"• {warning}")

    columns = st.columns(6)
    values = (
        ("PO 주문행", preflight["order_counts"]["PO"]),
        ("FOC 주문행", preflight["order_counts"]["FOC"]),
        ("전체 요청수량", preflight["total_request_quantity"]),
        ("GD 통과재고", preflight["usable_quantity_by_warehouse"]["GD"]),
        ("DU 통과재고", preflight["usable_quantity_by_warehouse"]["DU"]),
        ("사용 가능 LOT", preflight["grouped_lot_count"]),
    )
    for column, (label, value) in zip(columns, values):
        column.metric(label, _fmt_number(value))

    with st.expander("사전검증 상세", expanded=False):
        report = preflight["inventory_report"]
        detail_rows = [
            ("주문 파일", preflight["order_name"]),
            ("재고 파일", preflight["inventory_name"]),
            ("PO 요청수량", preflight["request_quantities"]["PO"]),
            ("FOC 요청수량", preflight["request_quantities"]["FOC"]),
            ("재고 없는 주문행", preflight["missing_inventory_rows"]),
            ("제품명 불일치 후보", preflight["product_mismatch_rows"]),
            ("상태 제외 수량", report.get("상태 제외 수량", 0)),
            ("GL 제외 수량", report.get("GL 제외 수량", 0)),
            ("EXP 미확인 수량", report.get("EXP 미확인 수량", 0)),
            ("LOT 미확인 수량", report.get("LOT 미확인 수량", 0)),
        ]
        detail_df = pd.DataFrame(detail_rows, columns=["검증 항목", "결과"])
        detail_df["결과"] = detail_df["결과"].map(
            lambda value: f"{value:,}" if isinstance(value, (int, float)) else str(value)
        )
        st.dataframe(detail_df, hide_index=True, width="stretch")


def _dataframe_xlsx_bytes(dataframe: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()


def _show_result(result_state: dict) -> None:
    metrics = result_state["metrics"]
    st.success("배정이 완료됐습니다. 검증 결과를 확인한 후 엑셀을 내려받으세요.")

    metric_columns = st.columns(5)
    metric_values = (
        ("전체 요청수량", metrics["total_request"]),
        ("전체 배정수량", metrics["total_allocated"]),
        ("미배정수량", metrics["unallocated"]),
        ("미배정 주문행", metrics["unallocated_lines"]),
        ("결과 상세행", metrics["result_rows"]),
    )
    for column, (label, value) in zip(metric_columns, metric_values):
        column.metric(label, _fmt_number(value))

    checks = result_state["integrity_checks"]
    warnings = [item for item in checks if item["상태"] == "WARNING"]
    if warnings:
        st.warning(
            " / ".join(f"{item['검증 항목']} {item['결과']}" for item in warnings)
            + " — '검증 요약'에서 확인해 주세요."
        )

    result_tab, summary_tab, validation_tab = st.tabs(
        ["LOT 배정 결과", "주문 요약", "검증 요약"]
    )
    with result_tab:
        result_df = pd.DataFrame(result_state["result_data"])
        if result_df.empty:
            st.info("표시할 배정 결과가 없습니다.")
        else:
            filter_columns = st.columns([1, 1, 1, 2])
            selected_status = filter_columns[0].selectbox(
                "배정상태", ["전체"] + sorted(result_df["배정상태"].dropna().astype(str).unique()), key="result_status_filter"
            )
            selected_type = filter_columns[1].selectbox(
                "구분", ["전체"] + sorted(result_df["구분"].dropna().astype(str).unique()), key="result_type_filter"
            )
            selected_warehouse = filter_columns[2].selectbox(
                "배정창고", ["전체"] + sorted(result_df["배정창고"].dropna().astype(str).unique()), key="result_warehouse_filter"
            )
            search_text = filter_columns[3].text_input(
                "검색", placeholder="바코드, 제품명, LOT 번호", key="result_search"
            )

            check_columns = st.columns(3)
            only_unknown_exp = check_columns[0].checkbox("EXP 미확인만", key="only_unknown_exp")
            only_unknown_lot = check_columns[1].checkbox("LOT 미확인만", key="only_unknown_lot")
            only_mismatch = check_columns[2].checkbox("제품명 불일치만", key="only_mismatch")

            filtered = result_df.copy()
            if selected_status != "전체":
                filtered = filtered[filtered["배정상태"].astype(str) == selected_status]
            if selected_type != "전체":
                filtered = filtered[filtered["구분"].astype(str) == selected_type]
            if selected_warehouse != "전체":
                filtered = filtered[filtered["배정창고"].astype(str) == selected_warehouse]
            if only_unknown_exp:
                filtered = filtered[filtered["유통기한 상태"] == "미확인"]
            if only_unknown_lot:
                filtered = filtered[filtered["배정된 로트번호"] == "미확인"]
            if only_mismatch:
                filtered = filtered[filtered["제품명 일치 여부"] == "불일치"]
            if search_text.strip():
                search_mask = pd.Series(False, index=filtered.index)
                for column in ("제품 바코드", "제품명", "배정된 로트번호"):
                    search_mask |= filtered[column].astype(str).str.contains(search_text.strip(), case=False, na=False, regex=False)
                filtered = filtered[search_mask]

            st.caption(f"조회 {len(filtered):,}행 / 전체 {len(result_df):,}행")
            preview = filtered.head(MAX_PREVIEW_ROWS)
            if len(filtered) > MAX_PREVIEW_ROWS:
                st.info(f"웹 화면에는 첫 {MAX_PREVIEW_ROWS:,}행만 표시하며 전체 결과는 엑셀에 포함됩니다.")
            st.dataframe(preview, hide_index=True, width="stretch", height=650)

    with summary_tab:
        summary_df = pd.DataFrame(result_state["summary_data"])
        if summary_df.empty:
            st.info("표시할 주문 요약이 없습니다.")
        else:
            summary_columns = st.columns([1, 1, 2])
            summary_status = summary_columns[0].selectbox(
                "주문 상태", ["전체"] + sorted(summary_df["상태"].dropna().astype(str).unique()), key="summary_status_filter"
            )
            summary_type = summary_columns[1].selectbox(
                "주문 구분", ["전체"] + sorted(summary_df["구분"].dropna().astype(str).unique()), key="summary_type_filter"
            )
            summary_search = summary_columns[2].text_input(
                "주문 검색", placeholder="바코드 또는 제품명", key="summary_search"
            )
            filtered_summary = summary_df.copy()
            if summary_status != "전체":
                filtered_summary = filtered_summary[filtered_summary["상태"].astype(str) == summary_status]
            if summary_type != "전체":
                filtered_summary = filtered_summary[filtered_summary["구분"].astype(str) == summary_type]
            if summary_search.strip():
                mask = (
                    filtered_summary["제품 바코드"].astype(str).str.contains(summary_search.strip(), case=False, na=False, regex=False)
                    | filtered_summary["제품명"].astype(str).str.contains(summary_search.strip(), case=False, na=False, regex=False)
                )
                filtered_summary = filtered_summary[mask]
            st.caption(f"조회 {len(filtered_summary):,}행 / 전체 {len(summary_df):,}행")
            st.dataframe(filtered_summary.head(MAX_PREVIEW_ROWS), hide_index=True, width="stretch", height=650)

            unallocated_df = summary_df[summary_df["상태"] != "배정완료"]
            if not unallocated_df.empty:
                st.download_button(
                    "미배정 주문만 다운로드",
                    data=_dataframe_xlsx_bytes(unallocated_df, "미배정 주문"),
                    file_name=f"미배정_주문_{result_state['file_timestamp']}.xlsx",
                    mime=XLSX_MIME, on_click="ignore",
                )

    with validation_tab:
        check_df = pd.DataFrame(checks)
        status_counts = check_df["상태"].value_counts().to_dict()
        st.caption(" / ".join(f"{status}: {status_counts.get(status, 0):,}건" for status in ("PASS", "WARNING", "FAIL")))
        st.dataframe(check_df, hide_index=True, width="stretch")

        with st.expander("입력·실행 정보", expanded=False):
            report = result_state["inventory_report"]
            information_rows = [
                ("앱 버전", APP_VERSION),
                ("실행 일시", result_state["executed_at"]),
                ("주문 파일", result_state["order_name"]),
                ("재고 파일", result_state["inventory_name"]),
                ("전체 재고행", report.get("전체 재고행", 0)),
                ("통과 재고수량", report.get("통과 재고수량", 0)),
                ("상태 제외 수량", report.get("상태 제외 수량", 0)),
                ("GL 제외 수량", report.get("GL 제외 수량", 0)),
            ]
            information_df = pd.DataFrame(information_rows, columns=["항목", "결과"])
            information_df["결과"] = information_df["결과"].map(
                lambda value: f"{value:,}" if isinstance(value, (int, float)) else str(value)
            )
            st.dataframe(information_df, hide_index=True, width="stretch")

    download_column, clear_column = st.columns([4, 1])
    download_column.download_button(
        label="전체 결과 엑셀 다운로드",
        data=result_state["output_bytes"], file_name=result_state["download_name"],
        mime=XLSX_MIME, type="primary", width="stretch", on_click="ignore",
    )
    if clear_column.button("결과 초기화", width="stretch"):
        st.session_state.pop("allocation_result", None)
        st.rerun()


st.title("PO·FOC LOT 번호 자동 배정")
st.caption(f"주문·재고 파일 사전검증 후 LOT 배정 결과를 생성합니다. 버전 {APP_VERSION}")

rule_columns = st.columns(4)
with rule_columns[0]:
    _rule_card("재고 사용", "판정이 <b>통과</b>인 GD·DU 재고만 사용하며 GL은 제외합니다.")
with rule_columns[1]:
    _rule_card("기본 우선순위", "<b>PO → FOC</b>, <b>GD → DU</b> 순으로 배정합니다.")
with rule_columns[2]:
    _rule_card("LOT 순서", "빠른 EXP → 잔여수량이 적은 LOT → LOT 번호순입니다.")
with rule_columns[3]:
    _rule_card("재고 부족", "허용된 창고를 합쳐도 부족하면 해당 주문은 <b>전체 미배정</b>합니다.")

st.write("")
upload_columns = st.columns(2)
order_upload = upload_columns[0].file_uploader(
    "1. PO/FOC 주문 파일", type=["xlsx"],
    help="PO 또는 FOC 시트와 Barcode, Product, Quantity 열이 필요합니다.", key="order_upload",
)
inventory_upload = upload_columns[1].file_uploader(
    "2. 글로벌 재고 파일", type=["xlsx"],
    help="'글로벌' 또는 '글로벌 유통기한' 시트가 필요합니다.", key="inventory_upload",
)

input_signature = None
preflight = None
if order_upload is not None and inventory_upload is not None:
    input_signature = _input_signature(order_upload, inventory_upload)
    stored_preflight = st.session_state.get("preflight_result")
    if not stored_preflight or stored_preflight.get("signature") != input_signature:
        st.session_state.pop("allocation_result", None)
        try:
            with st.spinner("입력 파일을 사전검증하는 중입니다..."):
                st.session_state["preflight_result"] = _run_preflight(order_upload, inventory_upload, input_signature)
        except (InputFormatError, ValueError, OSError) as exc:
            st.session_state["preflight_result"] = {"signature": input_signature, "status": "FAIL", "error": str(exc)}
        except Exception:
            logging.exception("Preflight validation failed")
            st.session_state["preflight_result"] = {
                "signature": input_signature, "status": "FAIL",
                "error": "사전검증 중 예상하지 못한 오류가 발생했습니다.",
            }
    preflight = st.session_state.get("preflight_result")
    if preflight and preflight.get("status") == "FAIL":
        st.error(preflight["error"])
    elif preflight:
        _show_preflight(preflight)
elif order_upload is not None or inventory_upload is not None:
    st.session_state.pop("allocation_result", None)
    st.session_state.pop("preflight_result", None)
    st.info("나머지 파일도 업로드하면 자동으로 사전검증합니다.")
else:
    st.session_state.pop("allocation_result", None)
    st.session_state.pop("preflight_result", None)

with st.expander("고급 배정 설정", expanded=False):
    setting_columns = st.columns(3)
    first_warehouse = setting_columns[0].selectbox("1순위 창고", ["GD", "DU"], index=0)
    first_order_type = setting_columns[1].selectbox("1순위 주문", ["PO", "FOC"], index=0)
    allow_split = setting_columns[2].checkbox(
        "단일 창고 부족 시 창고 합산 허용", value=True,
        help="두 창고 모두 단독 전량 배정이 불가능하지만 합산 재고가 충분할 때만 분할 배정합니다.",
    )
    st.caption("통과 재고만 사용, GL 제외, EXP 미확인 재고 후순위 배정 규칙은 고정됩니다.")

second_warehouse = "DU" if first_warehouse == "GD" else "GD"
second_order_type = "FOC" if first_order_type == "PO" else "PO"
warehouse_priority = (first_warehouse, second_warehouse)
order_priority = (first_order_type, second_order_type)

if warehouse_priority != ("GD", "DU") or order_priority != ("PO", "FOC") or not allow_split:
    st.warning(
        f"기본값과 다른 설정입니다: {' → '.join(order_priority)} / "
        f"{' → '.join(warehouse_priority)} / 창고 합산 {'허용' if allow_split else '비허용'}"
    )

preflight_valid = bool(preflight and preflight.get("status") in {"PASS", "WARNING"})
submitted = st.button("LOT 배정 실행", type="primary", width="stretch", disabled=not preflight_valid)

current_run_signature = None
if input_signature:
    current_run_signature = _run_signature(input_signature, warehouse_priority, allow_split, order_priority)

if submitted:
    st.session_state.pop("allocation_result", None)
    try:
        order_bytes = _validate_upload(order_upload, "PO/FOC 주문 파일")
        inventory_bytes = _validate_upload(inventory_upload, "글로벌 재고 파일")
        now = datetime.now(KOREA_TZ)
        with st.spinner("재고를 검산하고 LOT를 배정하는 중입니다..."):
            with tempfile.TemporaryDirectory(prefix="lot_allocation_") as temp_directory:
                root = Path(temp_directory)
                order_path = root / "orders.xlsx"
                inventory_path = root / "inventory.xlsx"
                output_path = root / "PO_FOC_LOT_배정_결과.xlsx"
                order_path.write_bytes(order_bytes)
                inventory_path.write_bytes(inventory_bytes)
                result = create_allocation_file(
                    order_path, inventory_path, output_path,
                    warehouse_priority=warehouse_priority, allow_split=allow_split,
                    order_priority=order_priority,
                    order_display_name=Path(order_upload.name).name,
                    inventory_display_name=Path(inventory_upload.name).name,
                    executed_at=now,
                )
                output_bytes = output_path.read_bytes()
        del order_bytes, inventory_bytes

        st.session_state["allocation_result"] = {
            "signature": current_run_signature,
            "metrics": {
                "total_request": result["total_request"], "total_allocated": result["total_allocated"],
                "unallocated": result["unallocated"], "unallocated_lines": result["shortage_lines"],
                "result_rows": result["result_rows"],
            },
            "result_data": result["result_data"], "summary_data": result["summary_data"],
            "inventory_report": result["inventory_report"], "integrity_checks": result["integrity_checks"],
            "output_bytes": output_bytes, "download_name": f"PO_FOC_LOT배정_{now:%Y%m%d_%H%M%S}.xlsx",
            "file_timestamp": f"{now:%Y%m%d_%H%M%S}", "executed_at": f"{now:%Y-%m-%d %H:%M:%S KST}",
            "order_name": Path(order_upload.name).name, "inventory_name": Path(inventory_upload.name).name,
        }
    except (InputFormatError, ValueError, OSError, RuntimeError) as exc:
        st.error(str(exc))
    except Exception:
        logging.exception("LOT allocation failed")
        st.error("처리 중 예상하지 못한 오류가 발생했습니다. 입력 파일을 확인해 주세요.")

result_state = st.session_state.get("allocation_result")
if result_state and current_run_signature == result_state.get("signature"):
    st.divider()
    _show_result(result_state)
elif result_state:
    st.info("파일 또는 배정 설정이 변경됐습니다. 새 설정으로 다시 실행해 주세요.")

st.divider()
st.markdown(
    '<div class="muted">업로드한 파일은 배정 처리용 임시 경로에서만 사용되며 '
    '처리 완료 후 제거됩니다. 회사 정보가 포함된 파일은 비공개 앱에서만 사용하세요.</div>',
    unsafe_allow_html=True,
)
