from __future__ import annotations

import hashlib
import logging
import tempfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from PO_FOC_LOT_배정_자동화 import InputFormatError, create_allocation_file


MAX_UPLOAD_BYTES = 100 * 1024 * 1024
KOREA_TZ = ZoneInfo("Asia/Seoul")


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
        background: #f7f9fc;
        border: 1px solid #e1e7ef;
        border-radius: 12px;
        padding: 14px 16px;
      }
      .rule-card {
        background: #f7f9fc;
        border: 1px solid #e1e7ef;
        border-radius: 12px;
        padding: 14px 16px;
        height: 100%;
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
        raise InputFormatError(f"{label}은 100MB보다 클 수 없습니다.")
    return payload


def _input_signature(
    order_bytes: bytes,
    inventory_bytes: bytes,
    warehouse_priority: tuple[str, str],
    allow_split: bool,
    order_priority: tuple[str, str],
) -> str:
    digest = hashlib.sha256()
    digest.update(order_bytes)
    digest.update(inventory_bytes)
    digest.update("|".join(warehouse_priority).encode("utf-8"))
    digest.update(str(allow_split).encode("ascii"))
    digest.update("|".join(order_priority).encode("utf-8"))
    return digest.hexdigest()


def _rule_card(title: str, text: str) -> None:
    st.markdown(
        f'<div class="rule-card"><div class="rule-title">{title}</div>'
        f'<div class="rule-text">{text}</div></div>',
        unsafe_allow_html=True,
    )


def _show_result(result_state: dict) -> None:
    metrics = result_state["metrics"]
    st.success("배정이 완료됐습니다. 아래 요약과 검증 결과를 확인한 후 엑셀을 내려받으세요.")

    metric_columns = st.columns(5)
    metric_values = (
        ("전체 요청수량", metrics["total_request"]),
        ("전체 배정수량", metrics["total_allocated"]),
        ("미배정수량", metrics["unallocated"]),
        ("미배정 주문행", metrics["shortage_lines"]),
        ("결과 상세행", metrics["result_rows"]),
    )
    for column, (label, value) in zip(metric_columns, metric_values):
        column.metric(label, _fmt_number(value))

    result_rows = result_state["result_data"]
    summaries = result_state["summary_data"]
    report = result_state["inventory_report"]

    mismatch_rows = sum(
        1 for row in result_rows if row.get("제품명 일치 여부") == "불일치"
    )
    unknown_exp_allocated = sum(
        int(row.get("LOT 배정수량") or 0)
        for row in result_rows
        if row.get("유통기한 상태") == "미확인"
    )
    unknown_lot_allocated = sum(
        int(row.get("LOT 배정수량") or 0)
        for row in result_rows
        if row.get("배정된 로트번호") == "미확인"
    )

    warning_messages = []
    if metrics["shortage_lines"]:
        warning_messages.append(f"미배정 주문 {int(metrics['shortage_lines']):,}행")
    if mismatch_rows:
        warning_messages.append(f"제품명 불일치 상세 {mismatch_rows:,}행")
    if unknown_exp_allocated:
        warning_messages.append(f"EXP 미확인 배정 {unknown_exp_allocated:,}개")
    if unknown_lot_allocated:
        warning_messages.append(f"LOT 미확인 배정 {unknown_lot_allocated:,}개")
    if warning_messages:
        st.warning(" / ".join(warning_messages) + " — '검증 요약' 탭에서 확인해 주세요.")

    result_tab, summary_tab, validation_tab = st.tabs(
        ["LOT 배정 결과", "주문 요약", "검증 요약"]
    )
    with result_tab:
        result_df = pd.DataFrame(result_rows)
        if result_df.empty:
            st.info("표시할 배정 결과가 없습니다.")
        else:
            st.dataframe(
                result_df,
                hide_index=True,
                width="stretch",
                height=min(720, 78 + 35 * len(result_df)),
            )

    with summary_tab:
        summary_df = pd.DataFrame(summaries)
        if summary_df.empty:
            st.info("표시할 주문 요약이 없습니다.")
        else:
            status_counts = summary_df["상태"].value_counts().to_dict()
            st.caption(
                " / ".join(f"{status}: {count:,}행" for status, count in status_counts.items())
            )
            st.dataframe(
                summary_df,
                hide_index=True,
                width="stretch",
                height=min(720, 78 + 35 * len(summary_df)),
            )

    with validation_tab:
        validation_rows = [
            ("실행 일시", result_state["executed_at"]),
            ("주문 파일", result_state["order_name"]),
            ("재고 파일", result_state["inventory_name"]),
            ("전체 주문행", len(summaries)),
            ("배정완료 주문행", sum(1 for item in summaries if item["상태"] == "배정완료")),
            ("미배정 주문행", metrics["shortage_lines"]),
            ("전체 재고행", report.get("전체 재고행", 0)),
            ("GD/DU 재고행", report.get("GD/DU 재고행", 0)),
            ("통과 재고수량", report.get("통과 재고수량", 0)),
            ("상태 제외 수량", report.get("상태 제외 수량", 0)),
            ("GL 제외 수량", report.get("GL 제외 수량", 0)),
            ("EXP 미확인 재고수량", report.get("EXP 미확인 수량", 0)),
            ("LOT 미확인 재고수량", report.get("LOT 미확인 수량", 0)),
            ("EXP 미확인 배정수량", unknown_exp_allocated),
            ("LOT 미확인 배정수량", unknown_lot_allocated),
            ("제품명 불일치 상세행", mismatch_rows),
        ]
        validation_df = pd.DataFrame(validation_rows, columns=["검증 항목", "결과"])
        validation_df["결과"] = validation_df["결과"].map(
            lambda value: f"{value:,}" if isinstance(value, (int, float)) else str(value)
        )
        st.dataframe(validation_df, hide_index=True, width="stretch")

    st.download_button(
        label="결과 엑셀 다운로드",
        data=result_state["output_bytes"],
        file_name=result_state["download_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
    )


st.title("PO·FOC LOT 번호 자동 배정")
st.caption("주문 파일과 글로벌 재고 파일을 업로드하면 배정 결과 엑셀을 생성합니다.")

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
with st.form("allocation_form", clear_on_submit=False):
    upload_columns = st.columns(2)
    with upload_columns[0]:
        order_upload = st.file_uploader(
            "1. PO/FOC 주문 파일",
            type=["xlsx"],
            help="PO 또는 FOC 시트와 Barcode, Product, Quantity 열이 필요합니다.",
        )
    with upload_columns[1]:
        inventory_upload = st.file_uploader(
            "2. 글로벌 재고 파일",
            type=["xlsx"],
            help="'글로벌' 또는 '글로벌 유통기한' 시트가 필요합니다.",
        )

    with st.expander("고급 배정 설정", expanded=False):
        setting_columns = st.columns(3)
        with setting_columns[0]:
            first_warehouse = st.selectbox(
                "1순위 창고",
                options=["GD", "DU"],
                index=0,
            )
        with setting_columns[1]:
            first_order_type = st.selectbox(
                "1순위 주문",
                options=["PO", "FOC"],
                index=0,
            )
        with setting_columns[2]:
            allow_split = st.checkbox(
                "단일 창고 부족 시 창고 합산 허용",
                value=True,
                help="두 창고 모두 단독 전량 배정이 불가능하지만 합산 재고가 충분할 때만 분할 배정합니다.",
            )
        st.caption(
            "통과 재고만 사용, GL 제외, EXP 미확인 재고 후순위 배정 규칙은 고정됩니다."
        )

    submitted = st.form_submit_button(
        "LOT 배정 실행",
        type="primary",
        width="stretch",
        disabled=order_upload is None or inventory_upload is None,
    )

second_warehouse = "DU" if first_warehouse == "GD" else "GD"
second_order_type = "FOC" if first_order_type == "PO" else "PO"
warehouse_priority = (first_warehouse, second_warehouse)
order_priority = (first_order_type, second_order_type)

current_signature = None
if order_upload is not None and inventory_upload is not None:
    try:
        current_signature = _input_signature(
            _file_bytes(order_upload),
            _file_bytes(inventory_upload),
            warehouse_priority,
            allow_split,
            order_priority,
        )
    except Exception:
        current_signature = None

if submitted:
    st.session_state.pop("allocation_result", None)
    try:
        order_bytes = _validate_upload(order_upload, "PO/FOC 주문 파일")
        inventory_bytes = _validate_upload(inventory_upload, "글로벌 재고 파일")
        signature = _input_signature(
            order_bytes,
            inventory_bytes,
            warehouse_priority,
            allow_split,
            order_priority,
        )

        with st.spinner("파일을 검증하고 LOT를 배정하는 중입니다..."):
            with tempfile.TemporaryDirectory(prefix="lot_allocation_") as temp_directory:
                temp_root = Path(temp_directory)
                order_path = temp_root / f"order_{Path(order_upload.name).name}"
                inventory_path = temp_root / f"inventory_{Path(inventory_upload.name).name}"
                output_path = temp_root / "PO_FOC_LOT_배정_결과.xlsx"
                order_path.write_bytes(order_bytes)
                inventory_path.write_bytes(inventory_bytes)

                result = create_allocation_file(
                    order_path,
                    inventory_path,
                    output_path,
                    warehouse_priority=warehouse_priority,
                    allow_split=allow_split,
                    order_priority=order_priority,
                )
                output_bytes = output_path.read_bytes()

        now = datetime.now(KOREA_TZ)
        st.session_state["allocation_result"] = {
            "signature": signature,
            "metrics": {
                "total_request": result["total_request"],
                "total_allocated": result["total_allocated"],
                "unallocated": result["unallocated"],
                "shortage_lines": result["shortage_lines"],
                "result_rows": result["result_rows"],
            },
            "result_data": result["result_data"],
            "summary_data": result["summary_data"],
            "inventory_report": result["inventory_report"],
            "output_bytes": output_bytes,
            "download_name": f"PO_FOC_LOT배정_{now:%Y%m%d_%H%M%S}.xlsx",
            "executed_at": f"{now:%Y-%m-%d %H:%M:%S}",
            "order_name": Path(order_upload.name).name,
            "inventory_name": Path(inventory_upload.name).name,
        }
    except (InputFormatError, ValueError, OSError) as exc:
        st.error(str(exc))
    except Exception:
        logging.exception("LOT allocation failed")
        st.error("처리 중 예상하지 못한 오류가 발생했습니다. 입력 파일을 확인한 후 다시 실행해 주세요.")

result_state = st.session_state.get("allocation_result")
if result_state and current_signature == result_state.get("signature"):
    st.divider()
    _show_result(result_state)
elif result_state:
    st.info("파일 또는 배정 설정이 변경됐습니다. 새 설정으로 다시 실행해 주세요.")

st.divider()
st.markdown(
    '<div class="muted">업로드한 파일은 배정 처리용 임시 경로에서만 사용되며, '
    '처리 완료 후 제거됩니다. 회사 정보가 포함된 파일은 비공개 배포 환경에서만 사용하세요.</div>',
    unsafe_allow_html=True,
)
