# -*- coding: utf-8 -*-
import io
import unittest
from datetime import date

from openpyxl.reader.excel import load_workbook

from xml2xlsx import xml2xlsx, XML2XLSXTarget, CellRef


class CellRefTest(unittest.TestCase):

    def setUp(self):
        self.target = XML2XLSXTarget()

    def test_unicode_same_worksheet(self):
        self.target.start(tag='sheet', attrib={'title': 'test1'})
        cell = CellRef(self.target, 0, 0)
        self.assertEquals(unicode(cell), u'A1')

    def test_unicode_different_worksheet(self):
        self.target.start(tag='sheet', attrib={'title': 'test1'})
        cell = CellRef(self.target, 0, 0)
        self.target.end(tag='sheet')
        self.target.start(tag='sheet', attrib={'title': 'test2'})
        self.assertEquals(unicode(cell), u'test1!A1')


class XML2XLSXTest(unittest.TestCase):
    def test_single_row(self):
        template = """
        <sheet title="test">
            <row>
                <cell>test cell</cell>
                <cell>test cell2</cell>
            </row>
        </sheet>

        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)

        self.assertEquals(len(wb.worksheets), 1,
                          u"Created workbook should have only one sheet")
        self.assertIn("test", wb.get_sheet_names(), u"Worksheet 'test' missing")
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A1"].value, u"test cell")
        self.assertEquals(ws["B1"].value, u"test cell2")

    def test_cell_font_format(self):
        template = """
        <sheet title="test">
            <row>
                <cell font="size: 10; bold: True;">test cell</cell>
            </row>
        </sheet>

        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)

        self.assertEquals(len(wb.worksheets), 1,
                          u"Created workbook should have only one sheet")
        self.assertIn("test", wb.get_sheet_names(), u"Worksheet 'test' missing")
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A1"].font.size, 10, "Font size not set properly")
        self.assertTrue(ws["A1"].font.bold, "Font is not bold")

    def test_unicode(self):
        template = """
        <sheet title="test">
            <row><cell>aąwźćńół</cell></row>
        </sheet>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A1"].value, u"aąwźćńół")

    def test_multiple_rows(self):
        template = """
        <sheet title="test">
            <row>
                <cell>test cell</cell>
            </row>
            <row>
                <cell>test cell2</cell>
            </row>
        </sheet>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A1"].value, u"test cell")
        self.assertEquals(ws["A2"].value, u"test cell2")

    def test_cell_type_number(self):
        template = u"""
        <sheet title="test"><row><cell type="number">1123.4</cell></row>
        </sheet>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A1"].value, 1123.4)

    def test_cell_type_date(self):
        template = u"""
        <sheet title="test">
            <row><cell type="date" date-fmt="%d.%m.%Y">24.01.1981</cell></row>
        </sheet>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A1"].value.date(), date(1981, 01, 24))

    def test_cell_ref_id(self):
        template = u"""
        <sheet title="test">
            <row><cell ref-id="refcell">XXXX</cell></row>
            <row><cell>{refcell}</cell></row>
        </sheet>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A2"].value, "A1")

    def test_cell_ref_id_different_worksheet(self):
        template = u"""
        <workbook>
            <sheet title="test">
                <row><cell ref-id="refcell">XXXX</cell></row>
                <row><cell>{refcell}</cell></row>
            </sheet>
            <sheet title="test2">
                <row><cell>{refcell}</cell></row>
            </sheet>
        </workbook>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        self.assertEquals(wb['test2']["A1"].value, "test!A1")

    def test_cell_ref_col(self):
        template = u"""
        <sheet title="test">
            <row><cell>{col}</cell><cell>{col}</cell></row>
        </sheet>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A1"].value, "1")
        self.assertEquals(ws["B1"].value, "2")

    def test_cell_ref_row(self):
        template = u"""
        <sheet title="test">
            <row><cell>{row}</cell></row>
            <row><cell>{row}</cell></row>
        </sheet>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A1"].value, "1")
        self.assertEquals(ws["A2"].value, "2")

    def test_cell_ref_append(self):
        template = u"""
        <sheet title="test">
            <row><cell ref-append="my-list">ABC</cell></row>
            <row><cell ref-append="my-list">DEFG</cell></row>
            <row><cell>{my-list}</cell></row>
        </sheet>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertEquals(ws["A3"].value, "A1, A2")


    def test_sheet_index_attrib(self):
        template = u"""
        <workbook>
            <sheet title="test">
            </sheet>
            <sheet title="test2" index="0">
            </sheet>
        </workbook>
        """
        sheet = io.BytesIO(xml2xlsx(template))
        wb = load_workbook(sheet)
        ws = wb.get_sheet_by_name("test")
        self.assertListEqual(wb.sheetnames, ["test2", "test"])


if __name__ == '__main__':
    unittest.main()
