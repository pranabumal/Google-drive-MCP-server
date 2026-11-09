import io
import unittest

from docx import Document
from openpyxl import Workbook

from src.server import _spreadsheet_text, _word_text


class ExtractorTests(unittest.TestCase):
    def test_extracts_values_from_multiple_spreadsheet_tabs(self):
        workbook = Workbook()
        first = workbook.active
        first.title = "Summary"
        first.append(["Name", "Amount"])
        first.append(["Example", 42])
        second = workbook.create_sheet("Notes")
        second.append(["Reviewed"])

        data = io.BytesIO()
        workbook.save(data)

        text = _spreadsheet_text(data.getvalue())

        self.assertIn("## Summary", text)
        self.assertIn("Name\tAmount", text)
        self.assertIn("Example\t42", text)
        self.assertIn("## Notes", text)
        self.assertIn("Reviewed", text)

    def test_extracts_word_paragraphs_and_tables(self):
        document = Document()
        document.add_paragraph("Document detail")
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Item"
        table.cell(0, 1).text = "Value"
        document.add_paragraph("After table")
        data = io.BytesIO()
        document.save(data)

        text = _word_text(data.getvalue())

        self.assertIn("Document detail", text)
        self.assertIn("Item\tValue", text)
        self.assertIn("After table", text)
        self.assertLess(text.index("Document detail"), text.index("Item\tValue"))
        self.assertLess(text.index("Item\tValue"), text.index("After table"))


if __name__ == "__main__":
    unittest.main()
