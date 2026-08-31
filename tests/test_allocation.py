from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from PO_FOC_LOT_배정_자동화 import allocate_orders, create_allocation_file


def make_lot(
    warehouse: str,
    quantity: int,
    *,
    barcode: str = "8800000000001",
    expiry: date | None = date(2028, 1, 1),
    lot_number: str = "LOT001",
) -> dict:
    return {
        "바코드": barcode,
        "재고 제품명": "TEST PRODUCT",
        "재고 제품명들": {"TEST PRODUCT"},
        "창고": warehouse,
        "유통기한": expiry,
        "유통기한 상태": "정상" if expiry else "미확인",
        "로트번호": lot_number,
        "원본행": [2],
        "원재고": quantity,
        "잔여재고": quantity,
    }


def make_order(quantity: int, order_type: str = "PO", row: int = 2) -> dict:
    return {
        "구분": order_type,
        "주문행": row,
        "바코드": "8800000000001",
        "제품명": "TEST PRODUCT",
        "요청수량": quantity,
    }


class AllocationLogicTests(unittest.TestCase):
    def test_gd_is_used_when_gd_can_fulfill(self):
        rows, summaries = allocate_orders(
            {"PO": [make_order(8)], "FOC": []},
            [make_lot("GD", 10), make_lot("DU", 20, lot_number="LOT002")],
        )
        self.assertEqual(rows[0]["배정창고"], "GD")
        self.assertEqual(summaries[0]["배정수량"], 8)

    def test_du_is_used_when_only_du_can_fulfill(self):
        rows, _ = allocate_orders(
            {"PO": [make_order(8)], "FOC": []},
            [make_lot("GD", 5), make_lot("DU", 10, lot_number="LOT002")],
        )
        self.assertEqual({row["배정창고"] for row in rows}, {"DU"})

    def test_split_uses_priority_order(self):
        rows, summaries = allocate_orders(
            {"PO": [make_order(10)], "FOC": []},
            [make_lot("GD", 6), make_lot("DU", 5, lot_number="LOT002")],
        )
        self.assertEqual([row["배정창고"] for row in rows], ["GD", "DU"])
        self.assertEqual([row["LOT 배정수량"] for row in rows], [6, 4])
        self.assertEqual(summaries[0]["상태"], "배정완료")

    def test_insufficient_order_does_not_consume_inventory(self):
        orders = {
            "PO": [make_order(11, row=2), make_order(10, row=3)],
            "FOC": [],
        }
        rows, summaries = allocate_orders(
            orders,
            [make_lot("GD", 6), make_lot("DU", 4, lot_number="LOT002")],
        )
        self.assertEqual(summaries[0]["상태"], "재고부족")
        self.assertEqual(summaries[0]["배정수량"], 0)
        self.assertEqual(summaries[1]["배정수량"], 10)
        self.assertEqual(sum(row["LOT 배정수량"] for row in rows), 10)

    def test_split_can_be_disabled_without_consuming_inventory(self):
        rows, summaries = allocate_orders(
            {"PO": [make_order(10)], "FOC": []},
            [make_lot("GD", 6), make_lot("DU", 5, lot_number="LOT002")],
            allow_split=False,
        )
        self.assertEqual(rows[0]["배정상태"], "분할불가")
        self.assertEqual(summaries[0]["배정수량"], 0)

    def test_du_priority_is_configurable(self):
        rows, _ = allocate_orders(
            {"PO": [make_order(5)], "FOC": []},
            [make_lot("GD", 10), make_lot("DU", 10, lot_number="LOT002")],
            warehouse_priority=("DU", "GD"),
        )
        self.assertEqual(rows[0]["배정창고"], "DU")

    def test_order_priority_is_configurable(self):
        orders = {
            "PO": [make_order(6, "PO", 2)],
            "FOC": [make_order(6, "FOC", 2)],
        }
        _, summaries = allocate_orders(
            orders,
            [make_lot("GD", 6)],
            order_priority=("FOC", "PO"),
        )
        self.assertEqual(summaries[0]["구분"], "FOC")
        self.assertEqual(summaries[0]["상태"], "배정완료")
        self.assertEqual(summaries[1]["상태"], "재고부족")

    def test_unknown_expiry_is_allocated_after_known_expiry(self):
        rows, _ = allocate_orders(
            {"PO": [make_order(7)], "FOC": []},
            [
                make_lot("GD", 5, expiry=None, lot_number="UNKNOWN"),
                make_lot("GD", 5, expiry=date(2029, 1, 1), lot_number="KNOWN"),
            ],
        )
        self.assertEqual([row["배정된 로트번호"] for row in rows], ["KNOWN", "UNKNOWN"])


class WorkbookIntegrationTests(unittest.TestCase):
    def test_end_to_end_workbook_can_be_reopened(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            order_path = root / "orders.xlsx"
            inventory_path = root / "inventory.xlsx"
            output_path = root / "result.xlsx"

            order_wb = Workbook()
            po_ws = order_wb.active
            po_ws.title = "PO"
            po_ws.append(["Barcode", "Product", "Quantity"])
            po_ws.append(["8800000000001", "TEST PRODUCT", 5])
            foc_ws = order_wb.create_sheet("FOC")
            foc_ws.append(["Barcode", "Product", "Quantity"])
            order_wb.save(order_path)
            order_wb.close()

            inventory_wb = Workbook()
            inventory_ws = inventory_wb.active
            inventory_ws.title = "글로벌"
            inventory_ws.append(
                [
                    "저장 위치",
                    "자재",
                    "자재내역",
                    "가용재고(ASN)",
                    "유통기한",
                    "로트번호(생산)",
                    "판정",
                ]
            )
            inventory_ws.append(
                ["GD", "8800000000001", "TEST PRODUCT", 10, date(2028, 1, 1), "LOT001", "통과"]
            )
            inventory_wb.save(inventory_path)
            inventory_wb.close()

            result = create_allocation_file(order_path, inventory_path, output_path)
            self.assertEqual(result["total_allocated"], 5)
            self.assertTrue(output_path.exists())

            reopened = load_workbook(output_path, data_only=False)
            try:
                self.assertEqual(reopened.sheetnames, ["LOT 배정", "주문 요약", "검증 요약"])
                self.assertEqual(reopened["LOT 배정"]["F9"].value, "배정완료")
            finally:
                reopened.close()


if __name__ == "__main__":
    unittest.main()
