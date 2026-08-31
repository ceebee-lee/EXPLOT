from __future__ import annotations

import unittest
from datetime import date
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest


def workbook_bytes(kind: str) -> bytes:
    buffer = BytesIO()
    workbook = Workbook()
    worksheet = workbook.active
    if kind == "order":
        worksheet.title = "PO"
        worksheet.append(["Barcode", "Product", "Quantity"])
        worksheet.append(["8800000000001", "TEST PRODUCT", 5])
        foc = workbook.create_sheet("FOC")
        foc.append(["Barcode", "Product", "Quantity"])
    else:
        worksheet.title = "글로벌"
        worksheet.append(
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
        worksheet.append(
            ["GD", "8800000000001", "TEST PRODUCT", 10, date(2028, 1, 1), "LOT001", "통과"]
        )
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


class StreamlitSmokeTests(unittest.TestCase):
    def test_app_opens_without_exception(self):
        app_path = Path(__file__).resolve().parents[1] / "app.py"
        app = AppTest.from_file(str(app_path), default_timeout=20)
        app.run()
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.title[0].value, "PO·FOC LOT 번호 자동 배정")
        self.assertEqual(len(app.file_uploader), 2)
        self.assertEqual(len(app.button), 1)

    def test_uploaded_files_produce_download(self):
        app_path = Path(__file__).resolve().parents[1] / "app.py"
        app = AppTest.from_file(str(app_path), default_timeout=20).run()
        app.file_uploader[0].upload(
            "orders.xlsx",
            workbook_bytes("order"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ).run()
        app.file_uploader[1].upload(
            "inventory.xlsx",
            workbook_bytes("inventory"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ).run()
        app.button[0].click().run()

        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.success), 1)
        self.assertEqual(len(app.metric), 5)
        self.assertEqual(len(app.download_button), 1)


if __name__ == "__main__":
    unittest.main()
