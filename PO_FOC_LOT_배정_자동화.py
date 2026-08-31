from __future__ import annotations

import argparse
import math
import os
import re
import sys
import tempfile
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from zoneinfo import ZoneInfo

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.datetime import from_excel
except ImportError as exc:
    raise SystemExit(
        "openpyxl이 설치되어 있지 않습니다. PowerShell에서 다음 명령을 먼저 실행하세요:\n"
        "python -m pip install openpyxl"
    ) from exc


ORDER_TYPES = ("PO", "FOC")
WAREHOUSE_PRIORITY = ("GD", "DU")
PASS_STATUS = "통과"
UNKNOWN_VALUE = "미확인"
KOREA_TZ = ZoneInfo("Asia/Seoul")
APP_VERSION = "2.0.0"
OUTPUT_HEADERS = [
    "구분",
    "주문 원본행",
    "제품 바코드",
    "제품명",
    "요청수량",
    "배정상태",
    "배정창고",
    "유통기한",
    "배정된 로트번호",
    "LOT 배정수량",
    "배정 전 LOT재고",
    "배정 시점 DU 가용재고",
    "배정 시점 GD 가용재고",
    "비고",
    "주문 ID",
    "재고 제품명",
    "제품명 일치 여부",
    "유통기한 상태",
    "재고 원본행",
]
ORDER_SUMMARY_HEADERS = [
    "주문 ID",
    "구분",
    "주문 원본행",
    "제품 바코드",
    "제품명",
    "요청수량",
    "배정수량",
    "미배정수량",
    "상태",
    "배정창고",
    "배정 시점 GD 가용재고",
    "배정 시점 DU 가용재고",
    "비고",
]


class InputFormatError(ValueError):
    """입력 파일의 시트 또는 필수 열 구조가 예상과 다를 때 발생한다."""


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).strip().lower()


def normalize_code(value: Any) -> str:
    """엑셀 숫자형 바코드의 .0을 제거하고 문자열로 통일한다."""
    if value in (None, ""):
        return ""
    if isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def normalize_lot(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def safe_number(value: Any) -> Optional[int]:
    """유한한 정수만 반환한다. NaN, 무한대, 소수, bool은 거부한다."""
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return None
    try:
        number = float(value) if isinstance(value, (int, float)) else float(
            str(value).replace(",", "").strip()
        )
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number) or not number.is_integer():
        return None
    return int(number)


def parse_date_value(value: Any) -> Tuple[Optional[date], str]:
    """날짜와 상태(정상/미확인/형식오류)를 함께 반환한다."""
    if value in (None, ""):
        return None, UNKNOWN_VALUE
    if isinstance(value, datetime):
        return value.date(), "정상"
    if isinstance(value, date):
        return value, "정상"
    if isinstance(value, (int, float)):
        try:
            converted = from_excel(value)
            parsed = converted.date() if isinstance(converted, datetime) else converted
            return parsed, "정상"
        except Exception:
            return None, "형식오류"

    text = str(value).strip()
    for fmt in (
        "%Y-%m-%d",
        "%Y.%m.%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%Y%m%d",
    ):
        try:
            return datetime.strptime(text, fmt).date(), "정상"
        except ValueError:
            continue
    return None, "형식오류"


def to_date_only(value: Any) -> Optional[date]:
    """기존 호출 호환용 날짜 변환 함수."""
    return parse_date_value(value)[0]


def is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def excel_safe_value(value: Any) -> Any:
    """입력 문자열이 결과 엑셀에서 수식으로 실행되지 않도록 한다."""
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def normalize_product_name(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).casefold()


def find_required_columns(ws, required_headers: Dict[str, Iterable[str]]) -> Dict[str, int]:
    """1행에서 헤더명을 찾아 0-based 열 인덱스를 반환한다."""
    header_map: Dict[str, List[int]] = defaultdict(list)
    for index, cell in enumerate(ws[1]):
        header_map[normalize_header(cell.value)].append(index)

    result: Dict[str, int] = {}
    missing: List[str] = []
    duplicated: List[str] = []

    for key, candidates in required_headers.items():
        candidate_list = list(candidates)
        matched_indexes: List[int] = []
        for candidate in candidate_list:
            matched_indexes.extend(header_map.get(normalize_header(candidate), []))
        matched_indexes = sorted(set(matched_indexes))

        if not matched_indexes:
            missing.append(" / ".join(candidate_list))
        elif len(matched_indexes) > 1:
            duplicated.append(" / ".join(candidate_list))
        else:
            result[key] = matched_indexes[0]

    if missing or duplicated:
        details: List[str] = []
        if missing:
            details.append(f"누락 열: {', '.join(missing)}")
        if duplicated:
            details.append(f"중복 열: {', '.join(duplicated)}")
        raise InputFormatError(
            f"'{ws.title}' 시트의 필수 열 구조가 올바르지 않습니다. " + " / ".join(details)
        )
    return result


def load_order_rows(order_path: Path) -> Dict[str, List[dict]]:
    """주문 파일의 PO/FOC 시트를 읽는다."""
    wb = load_workbook(order_path, data_only=True, read_only=True)
    formula_wb = load_workbook(order_path, data_only=False, read_only=True)
    try:
        result: Dict[str, List[dict]] = {order_type: [] for order_type in ORDER_TYPES}
        found_sheet = False

        for order_type in ORDER_TYPES:
            if order_type not in wb.sheetnames:
                continue
            found_sheet = True
            ws = wb[order_type]
            formula_ws = formula_wb[order_type]
            columns = find_required_columns(
                ws,
                {
                    "barcode": ("Barcode", "바코드"),
                    "product": ("Product", "제품명", "상품명"),
                    "quantity": ("Quantity", "수량", "주문수량"),
                },
            )

            cached_rows = ws.iter_rows(min_row=2, values_only=True)
            formula_rows = formula_ws.iter_rows(min_row=2, values_only=True)
            for excel_row, (row, formula_row) in enumerate(
                zip(cached_rows, formula_rows), start=2
            ):
                for field in ("barcode", "product", "quantity"):
                    index = columns[field]
                    cached_value = row[index] if len(row) > index else None
                    formula_value = formula_row[index] if len(formula_row) > index else None
                    if cached_value in (None, "") and is_formula(formula_value):
                        raise InputFormatError(
                            f"주문 파일 '{order_type}' 시트 {excel_row}행의 "
                            f"{field} 수식에 저장된 계산값이 없습니다. "
                            "Excel에서 수식을 계산한 뒤 저장하고 다시 실행해 주세요."
                        )

                barcode = normalize_code(row[columns["barcode"]] if len(row) > columns["barcode"] else None)
                product = row[columns["product"]] if len(row) > columns["product"] else None
                raw_quantity = row[columns["quantity"]] if len(row) > columns["quantity"] else None

                if not barcode and product in (None, "") and raw_quantity in (None, ""):
                    continue
                if not barcode:
                    raise InputFormatError(
                        f"주문 파일 '{order_type}' 시트 {excel_row}행에 Barcode가 없습니다."
                    )

                quantity = safe_number(raw_quantity)
                if quantity is None or quantity <= 0:
                    raise InputFormatError(
                        f"주문 파일 '{order_type}' 시트 {excel_row}행의 Quantity가 "
                        f"올바른 양의 정수가 아닙니다: {raw_quantity}"
                    )

                result[order_type].append(
                    {
                        "구분": order_type,
                        "주문행": excel_row,
                        "바코드": barcode,
                        "제품명": str(product or ""),
                        "요청수량": quantity,
                    }
                )

        if not found_sheet:
            raise InputFormatError("주문 파일에 'PO' 또는 'FOC' 시트가 없습니다.")
        if not any(result.values()):
            raise InputFormatError("주문 파일에 배정할 주문 데이터가 없습니다.")
        return result
    finally:
        wb.close()
        formula_wb.close()


def find_inventory_sheet(wb):
    for sheet_name in ("글로벌", "글로벌 유통기한"):
        if sheet_name in wb.sheetnames:
            return wb[sheet_name]
    raise InputFormatError("글로벌재고 파일에 '글로벌' 시트가 없습니다.")


def load_inventory_lots(inventory_path: Path) -> List[dict]:
    """기존 호출 호환용: 사용 가능한 LOT 목록만 반환한다."""
    lots, _ = load_inventory_lots_with_report(inventory_path)
    return lots


def load_inventory_lots_with_report(inventory_path: Path) -> Tuple[List[dict], dict]:
    """
    GD/DU 중 판정이 정확히 '통과'인 양의 정수 재고만 읽는다.

    같은 바코드·창고·EXP·LOT는 합산하되 원본행과 제품명 후보를 보존한다.
    실제 EXP 미확인은 유지하고, 날짜 형식 오류와 미계산 수식은 중단한다.
    """
    wb = load_workbook(inventory_path, data_only=True, read_only=True)
    formula_wb = load_workbook(inventory_path, data_only=False, read_only=True)
    try:
        ws = find_inventory_sheet(wb)
        formula_ws = formula_wb[ws.title]
        columns = find_required_columns(
            ws,
            {
                "location": ("저장 위치", "저장위치", "창고"),
                "material": ("자재", "Barcode", "바코드"),
                "product": ("자재내역", "제품명", "Product"),
                "quantity": ("가용재고(ASN)", "가용재고", "수량"),
                "expiry": ("유통기한", "EXP"),
                "lot": ("로트번호(생산)", "로트번호", "LOT"),
                "status": ("판정", "상태"),
            },
        )

        grouped: Dict[Tuple[str, str, Optional[date], str], dict] = {}
        report = {
            "전체 재고행": 0,
            "GD/DU 재고행": 0,
            "통과 재고행": 0,
            "통과 재고수량": 0,
            "상태 제외 행": 0,
            "상태 제외 수량": 0,
            "GL 제외 행": 0,
            "GL 제외 수량": 0,
            "0 재고 제외 행": 0,
            "EXP 미확인 행": 0,
            "EXP 미확인 수량": 0,
            "LOT 미확인 행": 0,
            "LOT 미확인 수량": 0,
            "그룹화 LOT 수": 0,
        }
        status_seen = False

        cached_rows = ws.iter_rows(min_row=2, values_only=True)
        formula_rows = formula_ws.iter_rows(min_row=2, values_only=True)
        for excel_row, (row, formula_row) in enumerate(
            zip(cached_rows, formula_rows), start=2
        ):
            cached_blank = all(value in (None, "") for value in row)
            formula_blank = all(value in (None, "") for value in formula_row)
            if cached_blank and formula_blank:
                continue
            report["전체 재고행"] += 1

            location_index = columns["location"]
            location = str(row[location_index] or "").strip().upper()
            raw_quantity = row[columns["quantity"]]
            quantity_for_report = safe_number(raw_quantity)

            if location == "GL":
                report["GL 제외 행"] += 1
                if quantity_for_report is not None and quantity_for_report > 0:
                    report["GL 제외 수량"] += quantity_for_report
                continue
            if location not in WAREHOUSE_PRIORITY:
                continue

            report["GD/DU 재고행"] += 1
            for field in ("material", "product", "quantity", "expiry", "lot", "status"):
                index = columns[field]
                cached_value = row[index] if len(row) > index else None
                formula_value = formula_row[index] if len(formula_row) > index else None
                if cached_value in (None, "") and is_formula(formula_value):
                    raise InputFormatError(
                        f"재고 파일 '{ws.title}' 시트 {excel_row}행의 {field} 수식에 "
                        "저장된 계산값이 없습니다. Excel에서 수식을 계산한 뒤 저장하고 "
                        "다시 실행해 주세요."
                    )

            status = str(row[columns["status"]] or "").strip()
            if status:
                status_seen = True
            if status != PASS_STATUS:
                report["상태 제외 행"] += 1
                if quantity_for_report is not None and quantity_for_report > 0:
                    report["상태 제외 수량"] += quantity_for_report
                continue

            material = normalize_code(row[columns["material"]])
            if not material:
                raise InputFormatError(
                    f"재고 파일 '{ws.title}' 시트 {excel_row}행의 바코드가 비어 있거나 "
                    "올바르지 않습니다."
                )

            quantity = safe_number(raw_quantity)
            if quantity is None:
                raise InputFormatError(
                    f"재고 파일 '{ws.title}' 시트 {excel_row}행의 가용재고가 "
                    f"유한한 정수가 아닙니다: {raw_quantity}"
                )
            if quantity < 0:
                raise InputFormatError(
                    f"재고 파일 '{ws.title}' 시트 {excel_row}행의 가용재고가 음수입니다: "
                    f"{quantity}"
                )
            if quantity == 0:
                report["0 재고 제외 행"] += 1
                continue

            expiry, expiry_status = parse_date_value(row[columns["expiry"]])
            if expiry_status == "형식오류":
                raise InputFormatError(
                    f"재고 파일 '{ws.title}' 시트 {excel_row}행의 유통기한 형식이 "
                    f"올바르지 않습니다: {row[columns['expiry']]}"
                )
            lot = normalize_lot(row[columns["lot"]]) or UNKNOWN_VALUE
            product = str(row[columns["product"]] or "").strip()
            key = (material, location, expiry, lot)

            if key not in grouped:
                grouped[key] = {
                    "바코드": material,
                    "재고 제품명": product,
                    "재고 제품명들": set(),
                    "창고": location,
                    "유통기한": expiry,
                    "유통기한 상태": expiry_status,
                    "로트번호": lot,
                    "원본행": [],
                    "원재고": 0,
                    "잔여재고": 0,
                }
            if product:
                grouped[key]["재고 제품명들"].add(product)
            grouped[key]["원본행"].append(excel_row)
            grouped[key]["원재고"] += quantity
            grouped[key]["잔여재고"] += quantity

            report["통과 재고행"] += 1
            report["통과 재고수량"] += quantity
            if expiry_status == UNKNOWN_VALUE:
                report["EXP 미확인 행"] += 1
                report["EXP 미확인 수량"] += quantity
            if lot == UNKNOWN_VALUE:
                report["LOT 미확인 행"] += 1
                report["LOT 미확인 수량"] += quantity

        if report["GD/DU 재고행"] == 0:
            raise InputFormatError("글로벌재고 파일에 GD 또는 DU 재고 행이 없습니다.")
        if not status_seen:
            raise InputFormatError(
                "글로벌재고 파일의 GD/DU 재고에서 '판정' 값을 읽지 못했습니다. "
                "파일을 Excel에서 열어 수식을 계산한 뒤 저장하고 다시 실행해 주세요."
            )

        lots = list(grouped.values())
        for item in lots:
            names = sorted(item["재고 제품명들"])
            item["재고 제품명"] = " | ".join(names)
        report["그룹화 LOT 수"] = len(lots)
        return lots, report
    finally:
        wb.close()
        formula_wb.close()


def candidate_sort_key(lot: dict):
    """빠른 EXP → 현재 잔여수량이 적은 LOT → LOT 번호순."""
    return (
        lot["유통기한"] or date.max,
        lot["잔여재고"],
        lot["로트번호"],
    )


def available_quantity(lots: List[dict], material: str, location: str) -> int:
    return sum(
        int(lot["잔여재고"])
        for lot in lots
        if lot["바코드"] == material and lot["창고"] == location and lot["잔여재고"] > 0
    )


def product_match_status(order_product: str, inventory_names: Iterable[str]) -> str:
    order_name = normalize_product_name(order_product)
    names = [name for name in inventory_names if name]
    if not order_name:
        return "주문명 없음"
    if not names:
        return "재고명 없음"
    return "일치" if any(normalize_product_name(name) == order_name for name in names) else "불일치"


def inspect_inputs(order_path: Path, inventory_path: Path) -> dict:
    """배정 전 입력 파일의 구조와 기초 수량을 검증한다."""
    for label, path in (("PO/FOC 주문 파일", order_path), ("글로벌 재고 파일", inventory_path)):
        if not path.exists() or not path.is_file():
            raise InputFormatError(f"{label}을 찾을 수 없습니다: {path}")
        if path.suffix.lower() != ".xlsx":
            raise InputFormatError(f"{label}은 .xlsx 파일이어야 합니다.")

    order_rows = load_order_rows(order_path)
    lots, inventory_report = load_inventory_lots_with_report(inventory_path)

    order_counts = {
        order_type: len(order_rows.get(order_type, [])) for order_type in ORDER_TYPES
    }
    request_quantities = {
        order_type: sum(int(row["요청수량"]) for row in order_rows.get(order_type, []))
        for order_type in ORDER_TYPES
    }
    usable_quantity_by_warehouse = {
        warehouse: sum(
            int(lot["잔여재고"])
            for lot in lots
            if lot["창고"] == warehouse
        )
        for warehouse in ("GD", "DU")
    }

    inventory_barcodes = {lot["바코드"] for lot in lots}
    orders = [row for order_type in ORDER_TYPES for row in order_rows.get(order_type, [])]
    missing_inventory_rows = [
        row for row in orders if row["바코드"] not in inventory_barcodes
    ]

    names_by_barcode: Dict[str, set[str]] = defaultdict(set)
    for lot in lots:
        names_by_barcode[lot["바코드"]].update(lot.get("재고 제품명들", set()))
    mismatch_rows = [
        row
        for row in orders
        if row["바코드"] in inventory_barcodes
        and product_match_status(row["제품명"], names_by_barcode[row["바코드"]]) == "불일치"
    ]

    warnings: List[str] = []
    if missing_inventory_rows:
        warnings.append(f"사용 가능한 재고가 없는 주문 {len(missing_inventory_rows):,}행")
    if mismatch_rows:
        warnings.append(f"주문·재고 제품명 불일치 후보 {len(mismatch_rows):,}행")
    if inventory_report.get("EXP 미확인 수량", 0):
        warnings.append(
            f"EXP 미확인 재고 {inventory_report['EXP 미확인 수량']:,}개"
        )
    if inventory_report.get("LOT 미확인 수량", 0):
        warnings.append(
            f"LOT 미확인 재고 {inventory_report['LOT 미확인 수량']:,}개"
        )

    return {
        "status": "WARNING" if warnings else "PASS",
        "order_counts": order_counts,
        "request_quantities": request_quantities,
        "total_order_rows": len(orders),
        "total_request_quantity": sum(request_quantities.values()),
        "usable_quantity_by_warehouse": usable_quantity_by_warehouse,
        "usable_inventory_quantity": sum(usable_quantity_by_warehouse.values()),
        "grouped_lot_count": len(lots),
        "missing_inventory_rows": len(missing_inventory_rows),
        "product_mismatch_rows": len(mismatch_rows),
        "inventory_report": inventory_report,
        "warnings": warnings,
    }


def validate_allocation_integrity(
    order_rows: Dict[str, List[dict]],
    lots: List[dict],
    result_rows: List[dict],
    summaries: List[dict],
    *,
    order_priority: Tuple[str, str] = ORDER_TYPES,
) -> List[dict]:
    """배정 결과가 주문·재고 보존 규칙을 만족하는지 검산한다."""
    checks: List[dict] = []

    def add_check(name: str, passed: bool, detail: str) -> None:
        checks.append(
            {"검증 항목": name, "상태": "PASS" if passed else "FAIL", "결과": detail}
        )

    expected_order_count = sum(len(order_rows.get(kind, [])) for kind in ORDER_TYPES)
    add_check(
        "주문행 보존",
        len(summaries) == expected_order_count,
        f"입력 {expected_order_count:,}행 / 요약 {len(summaries):,}행",
    )

    unique_ids = {item["주문 ID"] for item in summaries}
    add_check(
        "주문 ID 유일성",
        len(unique_ids) == len(summaries),
        f"전체 {len(summaries):,}건 / 유일 {len(unique_ids):,}건",
    )

    partial_orders = [
        item
        for item in summaries
        if float(item["배정수량"]) not in (0.0, float(item["요청수량"]))
    ]
    add_check(
        "전량 배정 원칙",
        not partial_orders,
        f"부분 배정 {len(partial_orders):,}행",
    )

    detail_by_order: Dict[str, float] = defaultdict(float)
    for row in result_rows:
        detail_by_order[row["주문 ID"]] += float(row.get("LOT 배정수량") or 0)
    mismatched_order_totals = [
        item
        for item in summaries
        if detail_by_order[item["주문 ID"]] != float(item["배정수량"])
    ]
    add_check(
        "주문별 상세·요약 합계",
        not mismatched_order_totals,
        f"불일치 {len(mismatched_order_totals):,}행",
    )

    detail_total = sum(detail_by_order.values())
    summary_total = sum(float(item["배정수량"]) for item in summaries)
    add_check(
        "전체 배정수량 합계",
        detail_total == summary_total,
        f"상세 {number_for_excel(detail_total):,} / 요약 {number_for_excel(summary_total):,}",
    )

    consumed_inventory = sum(
        int(lot["원재고"]) - int(lot["잔여재고"]) for lot in lots
    )
    add_check(
        "재고 차감·배정 합계",
        consumed_inventory == summary_total,
        f"재고 차감 {consumed_inventory:,} / 배정 {number_for_excel(summary_total):,}",
    )

    negative_lots = [lot for lot in lots if int(lot["잔여재고"]) < 0]
    add_check("음수 재고 방지", not negative_lots, f"음수 LOT {len(negative_lots):,}건")

    gl_rows = [row for row in result_rows if row.get("배정창고") == "GL"]
    add_check("GL 배정 제외", not gl_rows, f"GL 배정 {len(gl_rows):,}행")

    actual_sequence = [item["구분"] for item in summaries]
    first, second = order_priority
    first_after_second = any(
        kind == first and second in actual_sequence[:index]
        for index, kind in enumerate(actual_sequence)
    )
    add_check(
        "주문 우선순위",
        not first_after_second,
        f"적용 순서 {first} → {second}",
    )

    warning_specs = [
        (
            "미배정 주문",
            sum(1 for item in summaries if item["상태"] != "배정완료"),
            "행",
        ),
        (
            "제품명 불일치",
            sum(1 for row in result_rows if row.get("제품명 일치 여부") == "불일치"),
            "행",
        ),
        (
            "EXP 미확인 배정",
            sum(
                int(row.get("LOT 배정수량") or 0)
                for row in result_rows
                if row.get("유통기한 상태") == UNKNOWN_VALUE
            ),
            "개",
        ),
        (
            "LOT 미확인 배정",
            sum(
                int(row.get("LOT 배정수량") or 0)
                for row in result_rows
                if row.get("배정된 로트번호") == UNKNOWN_VALUE
            ),
            "개",
        ),
    ]
    for name, count, unit in warning_specs:
        checks.append(
            {
                "검증 항목": name,
                "상태": "WARNING" if count else "PASS",
                "결과": f"{count:,}{unit}",
            }
        )
    return checks


def allocate_orders(
    order_rows: Dict[str, List[dict]],
    lots: List[dict],
    *,
    warehouse_priority: Tuple[str, str] = WAREHOUSE_PRIORITY,
    allow_split: bool = True,
    order_priority: Tuple[str, str] = ORDER_TYPES,
) -> Tuple[List[dict], List[dict]]:
    """
    PO를 먼저 배정하고, 차감 후 남은 재고로 FOC를 배정한다.

    제품별 창고 배정:
      1) GD에서 전량 배정 가능하면 GD에서 전량 배정한다.
      2) GD가 부족하고 DU에서 전량 배정 가능하면 DU에서 전량 배정한다.
      3) 두 창고 모두 단독 전량 배정이 불가능하지만 합산 재고가 충분하면
         GD를 먼저 배정하고 부족분을 DU에서 이어서 배정한다.
      4) GD+DU 합산 재고도 부족하면 어느 창고에서도 전혀 배정하지 않는다.

    각 창고 안에서는 빠른 EXP → 현재 잔여수량이 적은 LOT → LOT 번호순으로 배정한다.
    """
    warehouse_priority = tuple(str(item).upper() for item in warehouse_priority)
    order_priority = tuple(str(item).upper() for item in order_priority)
    if len(warehouse_priority) != 2 or set(warehouse_priority) != {"GD", "DU"}:
        raise ValueError("warehouse_priority는 GD와 DU를 한 번씩 포함해야 합니다.")
    if len(order_priority) != 2 or set(order_priority) != {"PO", "FOC"}:
        raise ValueError("order_priority는 PO와 FOC를 한 번씩 포함해야 합니다.")

    result_rows: List[dict] = []
    order_summaries: List[dict] = []
    inventory_index: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
    available_by_key: Dict[Tuple[str, str], int] = defaultdict(int)
    for lot in lots:
        key = (lot["바코드"], lot["창고"])
        inventory_index[key].append(lot)
        available_by_key[key] += int(lot["잔여재고"])

    for order_type in order_priority:
        for order in order_rows.get(order_type, []):
            barcode = order["바코드"]
            requested = int(order["요청수량"])
            order_id = f"{order_type}-{order['주문행']}"
            du_available = available_by_key[(barcode, "DU")]
            gd_available = available_by_key[(barcode, "GD")]
            total_available = du_available + gd_available

            first_location, second_location = warehouse_priority
            available_by_location = {"GD": gd_available, "DU": du_available}
            first_available = available_by_location[first_location]
            second_available = available_by_location[second_location]
            allocation_blocked = total_available < requested or (
                not allow_split
                and first_available < requested
                and second_available < requested
            )

            # 전량 충족이 불가능하면 해당 주문은 재고를 전혀 차감하지 않는다.
            if allocation_blocked:
                if total_available < requested:
                    allocation_status = "재고부족"
                    note = (
                        "GD+DU 합산 가용재고가 요청수량보다 부족하여 전체 미배정 "
                        f"(합산 가용재고 {number_for_excel(total_available):,})"
                    )
                else:
                    allocation_status = "분할불가"
                    note = (
                        "단일 창고에서 전량 배정할 수 없고 창고 합산 배정이 "
                        "비활성화되어 전체 미배정"
                    )
                result_rows.append(
                    {
                        "구분": order_type,
                        "주문 원본행": order["주문행"],
                        "제품 바코드": barcode,
                        "제품명": order["제품명"],
                        "요청수량": number_for_excel(requested),
                        "배정상태": allocation_status,
                        "배정창고": None,
                        "유통기한": None,
                        "배정된 로트번호": None,
                        "LOT 배정수량": 0,
                        "배정 전 LOT재고": None,
                        "배정 시점 DU 가용재고": number_for_excel(du_available),
                        "배정 시점 GD 가용재고": number_for_excel(gd_available),
                        "비고": note,
                        "주문 ID": order_id,
                        "재고 제품명": None,
                        "제품명 일치 여부": None,
                        "유통기한 상태": None,
                        "재고 원본행": None,
                    }
                )
                order_summaries.append(
                    {
                        "주문 ID": order_id,
                        "구분": order_type,
                        "주문 원본행": order["주문행"],
                        "제품 바코드": barcode,
                        "제품명": order["제품명"],
                        "요청수량": number_for_excel(requested),
                        "배정수량": 0,
                        "미배정수량": number_for_excel(requested),
                        "상태": allocation_status,
                        "배정창고": None,
                        "배정 시점 GD 가용재고": number_for_excel(gd_available),
                        "배정 시점 DU 가용재고": number_for_excel(du_available),
                        "비고": note,
                    }
                )
                continue

            if first_available >= requested:
                selected_locations = (first_location,)
                allocation_note = f"{first_location}에서 전량 배정"
            elif second_available >= requested:
                selected_locations = (second_location,)
                allocation_note = f"{second_location}에서 전량 배정"
            else:
                selected_locations = warehouse_priority
                allocation_note = (
                    f"{first_location} 우선 배정 후 부족분을 "
                    f"{second_location}에서 합산 배정"
                )

            remaining = requested
            allocated_total = 0
            allocated_locations: List[str] = []
            first_detail_row = True

            # 단독 전량 배정 가능한 창고를 우선 사용하고, 둘 다 불가능할 때만 합산한다.
            for location in selected_locations:
                if remaining <= 0:
                    break

                candidates = [
                    lot
                    for lot in inventory_index.get((barcode, location), [])
                    if lot["잔여재고"] > 0
                ]
                candidates.sort(key=candidate_sort_key)

                for lot in candidates:
                    if remaining <= 0:
                        break

                    before = int(lot["잔여재고"])
                    assigned = min(remaining, before)
                    if assigned <= 0:
                        continue

                    lot["잔여재고"] = before - assigned
                    available_by_key[(barcode, location)] -= assigned
                    remaining -= assigned
                    allocated_total += assigned
                    if location not in allocated_locations:
                        allocated_locations.append(location)

                    inventory_names = sorted(lot.get("재고 제품명들", set()))
                    if not inventory_names and lot.get("재고 제품명"):
                        inventory_names = [lot["재고 제품명"]]

                    result_rows.append(
                        {
                            "구분": order_type,
                            "주문 원본행": order["주문행"],
                            "제품 바코드": barcode,
                            "제품명": order["제품명"],
                            "요청수량": (
                                number_for_excel(requested) if first_detail_row else None
                            ),
                            "배정상태": "배정완료",
                            "배정창고": location,
                            "유통기한": lot["유통기한"],
                            "배정된 로트번호": lot["로트번호"],
                            "LOT 배정수량": number_for_excel(assigned),
                            "배정 전 LOT재고": number_for_excel(before),
                            "배정 시점 DU 가용재고": number_for_excel(du_available),
                            "배정 시점 GD 가용재고": number_for_excel(gd_available),
                            "비고": allocation_note,
                            "주문 ID": order_id,
                            "재고 제품명": " | ".join(inventory_names),
                            "제품명 일치 여부": product_match_status(
                                order["제품명"], inventory_names
                            ),
                            "유통기한 상태": lot.get("유통기한 상태", "정상"),
                            "재고 원본행": ", ".join(
                                str(row_number) for row_number in lot.get("원본행", [])
                            ),
                        }
                    )
                    first_detail_row = False

            # 사전 합산 검증을 통과했으므로 정상이라면 요청수량이 반드시 전량 배정된다.
            if remaining > 0:
                raise RuntimeError(
                    f"내부 배정 오류: {order_type} {barcode}의 GD+DU 합산 재고는 충분했지만 "
                    f"요청수량을 모두 배정하지 못했습니다. 남은 수량: {remaining}"
                )

            order_summaries.append(
                {
                    "주문 ID": order_id,
                    "구분": order_type,
                    "주문 원본행": order["주문행"],
                    "제품 바코드": barcode,
                    "제품명": order["제품명"],
                    "요청수량": number_for_excel(requested),
                    "배정수량": number_for_excel(allocated_total),
                    "미배정수량": 0,
                    "상태": "배정완료",
                    "배정창고": "+".join(allocated_locations),
                    "배정 시점 GD 가용재고": number_for_excel(gd_available),
                    "배정 시점 DU 가용재고": number_for_excel(du_available),
                    "비고": allocation_note,
                }
            )

    return result_rows, order_summaries

def number_for_excel(value: float | int) -> float | int:
    number = float(value)
    return int(number) if number.is_integer() else number


def build_output_workbook(
    order_path: Path,
    inventory_path: Path,
    result_rows: List[dict],
    summaries: List[dict],
    inventory_report: Optional[dict] = None,
    allocation_settings: Optional[dict] = None,
    integrity_checks: Optional[List[dict]] = None,
    order_display_name: Optional[str] = None,
    inventory_display_name: Optional[str] = None,
    executed_at: Optional[datetime] = None,
) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "LOT 배정"
    ws.sheet_view.showGridLines = False

    max_col = len(OUTPUT_HEADERS)
    last_col_letter = get_column_letter(max_col)

    # 제목
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = "PO·FOC LOT 번호 자동 배정 결과"
    ws["A1"].font = Font(name="맑은 고딕", bold=True, size=16, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="17365D")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    po_request = sum(float(item["요청수량"]) for item in summaries if item["구분"] == "PO")
    po_allocated = sum(float(item["배정수량"]) for item in summaries if item["구분"] == "PO")
    foc_request = sum(float(item["요청수량"]) for item in summaries if item["구분"] == "FOC")
    foc_allocated = sum(float(item["배정수량"]) for item in summaries if item["구분"] == "FOC")
    total_request = po_request + foc_request
    total_allocated = po_allocated + foc_allocated
    shortage_count = sum(1 for item in summaries if item["상태"] != "배정완료")

    settings = allocation_settings or {}
    warehouse_priority = tuple(settings.get("warehouse_priority", WAREHOUSE_PRIORITY))
    order_priority = tuple(settings.get("order_priority", ORDER_TYPES))
    allow_split = bool(settings.get("allow_split", True))
    first_warehouse, second_warehouse = warehouse_priority
    warehouse_rule = f"{first_warehouse} 단독 → {second_warehouse} 단독"
    if allow_split:
        warehouse_rule += f" → 필요 시 {first_warehouse}+{second_warehouse} 합산"

    display_order_name = order_display_name or order_path.name
    display_inventory_name = inventory_display_name or inventory_path.name
    run_at = executed_at or datetime.now(KOREA_TZ)

    info_rows = [
        [
            "입력 주문 파일", display_order_name,
            "입력 재고 파일", display_inventory_name,
            "배정 기준", warehouse_rule,
            "사용 재고", "판정='통과'",
        ],
        [
            "PO 요청", number_for_excel(po_request),
            "PO 배정", number_for_excel(po_allocated),
            "FOC 요청", number_for_excel(foc_request),
            "FOC 배정", number_for_excel(foc_allocated),
        ],
        [
            "전체 요청", number_for_excel(total_request),
            "전체 배정", number_for_excel(total_allocated),
            "미배정", number_for_excel(total_request - total_allocated),
            "미배정 주문행", shortage_count,
        ],
        [
            "LOT 우선순위", "EXP 오름차순 → 같은 EXP는 현재 잔여수량 오름차순 → LOT 번호순",
            "창고 분할", (
                f"단독 배정 불가 시에만 {first_warehouse} → {second_warehouse} 합산"
                if allow_split else "허용하지 않음"
            ),
            "GL", "제외",
            "배정 순서", " → ".join(order_priority),
        ],
    ]

    for row_index, row_values in enumerate(info_rows, start=3):
        for col_index, value in enumerate(row_values, start=1):
            ws.cell(row_index, col_index).value = value

    label_fill = PatternFill("solid", fgColor="D9EAF7")
    label_font = Font(name="맑은 고딕", bold=True, size=10, color="1F1F1F")
    info_font = Font(name="맑은 고딕", size=10)
    thin_blue = Side(style="thin", color="A6BFD5")
    info_border = Border(left=thin_blue, right=thin_blue, top=thin_blue, bottom=thin_blue)

    for row_index in range(3, 7):
        for col_index in range(1, 9):
            cell = ws.cell(row_index, col_index)
            cell.font = info_font
            cell.border = info_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col_index % 2 == 1:
                cell.fill = label_fill
                cell.font = label_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 표 헤더 및 데이터
    header_row = 8
    for col_index, header in enumerate(OUTPUT_HEADERS, start=1):
        ws.cell(header_row, col_index).value = header

    for row_index, item in enumerate(result_rows, start=header_row + 1):
        for col_index, header in enumerate(OUTPUT_HEADERS, start=1):
            ws.cell(row_index, col_index).value = excel_safe_value(item.get(header))

    data_last_row = header_row + len(result_rows)
    if not result_rows:
        data_last_row = header_row

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
    body_font = Font(name="맑은 고딕", size=10)
    thin_gray = Side(style="thin", color="D9E2F3")
    body_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = body_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 32

    for row in ws.iter_rows(min_row=header_row + 1, max_row=max(data_last_row, header_row + 1), max_col=max_col):
        for cell in row:
            cell.font = body_font
            cell.border = body_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    # 표시 형식
    for row_index in range(header_row + 1, data_last_row + 1):
        ws.cell(row_index, 3).number_format = "@"           # 바코드
        ws.cell(row_index, 5).number_format = "#,##0"       # 요청수량
        ws.cell(row_index, 8).number_format = "yyyy-mm-dd"  # 유통기한
        for col_index in (10, 11, 12, 13):
            ws.cell(row_index, col_index).number_format = "#,##0"

    widths = {
        "A": 8,
        "B": 11,
        "C": 18,
        "D": 44,
        "E": 12,
        "F": 12,
        "G": 10,
        "H": 13,  # 유통기한
        "I": 22,  # 로트번호
        "J": 14,
        "K": 17,
        "L": 20,
        "M": 20,
        "N": 48,
        "O": 13,
        "P": 42,
        "Q": 16,
        "R": 15,
        "S": 20,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.freeze_panes = "A9"

    # Excel 표가 자동필터를 포함하므로 ws.auto_filter는 별도로 설정하지 않는다.
    warning_fill = PatternFill("solid", fgColor="FCE4D6")
    if result_rows:
        table = Table(
            displayName="LotAllocationTable",
            ref=f"A{header_row}:{last_col_letter}{data_last_row}",
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)

        ws.conditional_formatting.add(
            f"F{header_row + 1}:F{data_last_row}",
            FormulaRule(
                formula=[f'$F{header_row + 1}<>"배정완료"'],
                fill=warning_fill,
            ),
        )

    # 주문당 한 행으로 구성된 집계용 요약 시트
    summary_ws = wb.create_sheet("주문 요약")
    summary_ws.sheet_view.showGridLines = False
    for col_index, header in enumerate(ORDER_SUMMARY_HEADERS, start=1):
        cell = summary_ws.cell(1, col_index, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = body_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_index, item in enumerate(summaries, start=2):
        for col_index, header in enumerate(ORDER_SUMMARY_HEADERS, start=1):
            cell = summary_ws.cell(row_index, col_index, excel_safe_value(item.get(header)))
            cell.font = body_font
            cell.border = body_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    summary_ws.freeze_panes = "A2"
    summary_widths = [13, 8, 12, 18, 42, 12, 12, 12, 12, 14, 22, 22, 48]
    for col_index, width in enumerate(summary_widths, start=1):
        summary_ws.column_dimensions[get_column_letter(col_index)].width = width
    for row_index in range(2, len(summaries) + 2):
        summary_ws.cell(row_index, 4).number_format = "@"
        for col_index in (6, 7, 8, 11, 12):
            summary_ws.cell(row_index, col_index).number_format = "#,##0"
    if summaries:
        summary_table = Table(
            displayName="OrderSummaryTable",
            ref=f"A1:{get_column_letter(len(ORDER_SUMMARY_HEADERS))}{len(summaries) + 1}",
        )
        summary_table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        summary_ws.add_table(summary_table)
        summary_ws.conditional_formatting.add(
            f"I2:I{len(summaries) + 1}",
            FormulaRule(formula=['$I2<>"배정완료"'], fill=warning_fill),
        )

    # 입력 및 배정 과정의 감사·검증용 요약 시트
    validation_ws = wb.create_sheet("검증 요약")
    validation_ws.sheet_view.showGridLines = False
    validation_ws["A1"] = "검증 항목"
    validation_ws["B1"] = "상태"
    validation_ws["C1"] = "결과"
    for cell in validation_ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = body_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    report = inventory_report or {}
    mismatch_rows = sum(
        1 for row in result_rows if row.get("제품명 일치 여부") == "불일치"
    )
    allocated_unknown_exp = sum(
        int(row.get("LOT 배정수량") or 0)
        for row in result_rows
        if row.get("유통기한 상태") == UNKNOWN_VALUE
    )
    allocated_unknown_lot = sum(
        int(row.get("LOT 배정수량") or 0)
        for row in result_rows
        if row.get("배정된 로트번호") == UNKNOWN_VALUE
    )
    validation_rows = [
        ("앱 버전", "INFO", APP_VERSION),
        ("실행 일시", "INFO", run_at.strftime("%Y-%m-%d %H:%M:%S KST")),
        ("주문 파일", "INFO", display_order_name),
        ("재고 파일", "INFO", display_inventory_name),
        ("전체 주문행", "INFO", len(summaries)),
        ("배정완료 주문행", "INFO", sum(1 for item in summaries if item["상태"] == "배정완료")),
        ("미배정 주문행", "INFO", shortage_count),
        ("전체 재고행", "INFO", report.get("전체 재고행", 0)),
        ("GD/DU 재고행", "INFO", report.get("GD/DU 재고행", 0)),
        ("통과 재고수량", "INFO", report.get("통과 재고수량", 0)),
        ("상태 제외 수량", "INFO", report.get("상태 제외 수량", 0)),
        ("GL 제외 수량", "INFO", report.get("GL 제외 수량", 0)),
        ("EXP 미확인 재고수량", "INFO", report.get("EXP 미확인 수량", 0)),
        ("LOT 미확인 재고수량", "INFO", report.get("LOT 미확인 수량", 0)),
        ("EXP 미확인 배정수량", "INFO", allocated_unknown_exp),
        ("LOT 미확인 배정수량", "INFO", allocated_unknown_lot),
        ("제품명 불일치 상세행", "INFO", mismatch_rows),
    ]
    validation_rows.extend(
        (item["검증 항목"], item["상태"], item["결과"])
        for item in (integrity_checks or [])
    )
    status_fills = {
        "PASS": PatternFill("solid", fgColor="E2F0D9"),
        "WARNING": PatternFill("solid", fgColor="FFF2CC"),
        "FAIL": PatternFill("solid", fgColor="FCE4D6"),
        "INFO": PatternFill("solid", fgColor="D9EAF7"),
    }
    for row_index, (label, status, value) in enumerate(validation_rows, start=2):
        label_cell = validation_ws.cell(row_index, 1, label)
        status_cell = validation_ws.cell(row_index, 2, status)
        value_cell = validation_ws.cell(row_index, 3, excel_safe_value(value))
        label_cell.fill = label_fill
        label_cell.font = label_font
        status_cell.fill = status_fills.get(status, status_fills["INFO"])
        status_cell.font = Font(name="맑은 고딕", bold=True, size=10)
        status_cell.alignment = Alignment(horizontal="center", vertical="center")
        for cell in (label_cell, status_cell, value_cell):
            cell.border = info_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    validation_ws.column_dimensions["A"].width = 30
    validation_ws.column_dimensions["B"].width = 13
    validation_ws.column_dimensions["C"].width = 58
    validation_ws.freeze_panes = "A2"
    return wb


def validate_paths(order_path: Path, inventory_path: Path, output_path: Path) -> None:
    for label, path in (("주문 파일", order_path), ("글로벌재고 파일", inventory_path)):
        if not path.exists():
            raise FileNotFoundError(f"{label}을 찾을 수 없습니다: {path}")
        if not path.is_file():
            raise ValueError(f"{label} 경로가 파일이 아닙니다: {path}")
        if path.suffix.lower() != ".xlsx":
            raise ValueError(f"{label}은 .xlsx 파일이어야 합니다: {path}")

    resolved_inputs = {order_path.resolve(), inventory_path.resolve()}
    if output_path.resolve() in resolved_inputs:
        raise ValueError("결과 파일은 입력 주문 파일 또는 글로벌재고 파일과 같은 경로일 수 없습니다.")
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError(f"결과 파일은 .xlsx 파일이어야 합니다: {output_path}")


def create_allocation_file(
    order_path: Path,
    inventory_path: Path,
    output_path: Path,
    *,
    warehouse_priority: Tuple[str, str] = WAREHOUSE_PRIORITY,
    allow_split: bool = True,
    order_priority: Tuple[str, str] = ORDER_TYPES,
    order_display_name: Optional[str] = None,
    inventory_display_name: Optional[str] = None,
    executed_at: Optional[datetime] = None,
) -> dict:
    validate_paths(order_path, inventory_path, output_path)

    order_rows = load_order_rows(order_path)
    lots, inventory_report = load_inventory_lots_with_report(inventory_path)
    result_rows, summaries = allocate_orders(
        order_rows,
        lots,
        warehouse_priority=warehouse_priority,
        allow_split=allow_split,
        order_priority=order_priority,
    )
    integrity_checks = validate_allocation_integrity(
        order_rows,
        lots,
        result_rows,
        summaries,
        order_priority=order_priority,
    )
    failed_checks = [item for item in integrity_checks if item["상태"] == "FAIL"]
    if failed_checks:
        failed_names = ", ".join(item["검증 항목"] for item in failed_checks)
        raise RuntimeError(f"배정 결과 내부 검증에 실패했습니다: {failed_names}")
    workbook = build_output_workbook(
        order_path,
        inventory_path,
        result_rows,
        summaries,
        inventory_report,
        allocation_settings={
            "warehouse_priority": warehouse_priority,
            "allow_split": allow_split,
            "order_priority": order_priority,
        },
        integrity_checks=integrity_checks,
        order_display_name=order_display_name,
        inventory_display_name=inventory_display_name,
        executed_at=executed_at,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=f".{output_path.stem}.",
        suffix=output_path.suffix,
        dir=output_path.parent,
        delete=False,
    ) as temp_file:
        temp_path = Path(temp_file.name)
    try:
        workbook.save(temp_path)
        os.replace(temp_path, output_path)
    finally:
        workbook.close()
        if temp_path.exists():
            temp_path.unlink()

    total_request = sum(float(item["요청수량"]) for item in summaries)
    total_allocated = sum(float(item["배정수량"]) for item in summaries)
    return {
        "output": output_path,
        "total_request": total_request,
        "total_allocated": total_allocated,
        "unallocated": total_request - total_allocated,
        "shortage_lines": sum(1 for item in summaries if item["상태"] != "배정완료"),
        "result_rows": len(result_rows),
        "result_data": result_rows,
        "summary_data": summaries,
        "inventory_report": inventory_report,
        "integrity_checks": integrity_checks,
    }


def choose_paths_with_gui() -> Tuple[Path, Path, Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError as exc:
        raise RuntimeError(
            "파일 선택 창을 사용할 수 없습니다. --order, --inventory, --out 옵션으로 실행해 주세요."
        ) from exc

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        order_file = filedialog.askopenfilename(
            title="PO/FOC 주문 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if not order_file:
            raise RuntimeError("주문 파일 선택이 취소되었습니다.")

        inventory_file = filedialog.askopenfilename(
            title="글로벌재고 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if not inventory_file:
            raise RuntimeError("글로벌재고 파일 선택이 취소되었습니다.")

        default_output = Path(order_file).with_name("PO_FOC_LOT번호_배정시트.xlsx")
        output_file = filedialog.asksaveasfilename(
            title="결과 파일 저장",
            defaultextension=".xlsx",
            initialdir=str(default_output.parent),
            initialfile=default_output.name,
            filetypes=[("Excel 파일", "*.xlsx")],
        )
    finally:
        root.destroy()

    if not output_file:
        raise RuntimeError("결과 파일 저장이 취소되었습니다.")
    return Path(order_file), Path(inventory_file), Path(output_file)


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "PO/FOC 주문 파일과 글로벌재고 파일을 이용해 GD 단독, DU 단독, GD+DU 합산 "
            "순서로 LOT를 배정합니다. 인자를 생략하면 파일 선택 창이 열립니다."
        )
    )
    parser.add_argument("--order", help="PO/FOC 주문 파일 경로")
    parser.add_argument("--inventory", help="글로벌재고 파일 경로")
    parser.add_argument("--out", help="결과 파일 경로")
    return parser.parse_args()


def main() -> None:
    args = parse_arguments()

    try:
        supplied = [args.order, args.inventory, args.out]
        if any(supplied) and not all(supplied):
            raise ValueError("명령어로 실행할 때는 --order, --inventory, --out을 모두 입력해야 합니다.")

        if all(supplied):
            order_path = Path(args.order)
            inventory_path = Path(args.inventory)
            output_path = Path(args.out)
        else:
            order_path, inventory_path, output_path = choose_paths_with_gui()

        result = create_allocation_file(order_path, inventory_path, output_path)
    except (OSError, InputFormatError, ValueError, RuntimeError) as exc:
        print(f"오류: {exc}", file=sys.stderr)
        if sys.stdin.isatty():
            input("\nEnter를 누르면 종료합니다...")
        raise SystemExit(1) from exc

    print(f"완료: {result['output']}")
    print(f"전체 요청수량: {number_for_excel(result['total_request']):,}")
    print(f"전체 배정수량: {number_for_excel(result['total_allocated']):,}")
    print(f"미배정수량: {number_for_excel(result['unallocated']):,}")
    print(f"미배정 주문행: {result['shortage_lines']:,}")
    print(f"결과 상세행: {result['result_rows']:,}")

    if sys.stdin.isatty() and not all([args.order, args.inventory, args.out]):
        input("\nEnter를 누르면 종료합니다...")


if __name__ == "__main__":
    main()
