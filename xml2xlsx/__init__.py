# -*- coding: utf-8 -*-
import logging

from lxml import etree

from openpyxl import Workbook
from openpyxl.writer.dump_worksheet import WriteOnlyCell
from openpyxl.writer.excel import save_virtual_workbook

logger = logging.getLogger(__name__)


class XML2XLSXTarget(object):
    def __init__(self):
        self.wb = Workbook(encoding='utf-8')
        self._current_ws = None
        self._row_buf = []
        self._cell = None

    def start(self, tag, attrib):
        if tag == 'sheet':
            if not self._current_ws:
                self._current_ws = self.wb.active
            else:
                self._current_ws = self.wb.create_sheet()
            if 'name' in attrib:
                self._current_ws.title = attrib['name']
        elif tag == 'row':
            self._row_buf = []
        elif tag == 'cell':
            self._cell = WriteOnlyCell(self._current_ws)

    def data(self, data):
        if self._cell:
            if self._cell.value:
                # TODO: Szybki fix na to, że znakiunicode powodują przerwanie
                #  czytania data i rozbijają to na 2
                self._cell.value += data
            else:
                self._cell.value = data

    def end(self, tag):
        if tag == 'sheet':
            pass
        elif tag == 'row':
            self._current_ws.append(self._row_buf)
            self._row_buf = []
        elif tag == 'cell':
            self._row_buf.append(self._cell)
            self._cell = None

    def close(self):
        return save_virtual_workbook(self.wb)


def xml2xlsx(xml):
    parser = etree.XMLParser(target=XML2XLSXTarget(), encoding='UTF-8',
                             remove_blank_text=True)
    return etree.XML(xml, parser, )
