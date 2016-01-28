# -*- coding: utf-8 -*-
import logging

from lxml import etree
from datetime import datetime

from decimal import Decimal

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
        self._cell_type = None
        self._cell_date_format = None


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
            ctype = attrib.get('type', 'unicode')
            if ctype not in ['unicode', 'number', 'date']:
                raise ValueError(u'Unknown cell type {ctype}.'.format(
                    ctype=ctype,
                ))
            self._cell_type = ctype
            try:
                self._cell_date_format = attrib.get('date-fmt')
            except KeyError:
                raise ValueError(u"Specify 'date-fmt' attribute for 'date'"
                                 u" type")

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
            if self._cell_type == 'number':
                self._cell.value = Decimal(self._cell.value)
            elif self._cell_type == 'date':
                self._cell.value = datetime.strptime(
                        self._cell.value, self._cell_date_format).date()
            self._row_buf.append(self._cell)
            self._cell = None

    def close(self):
        return save_virtual_workbook(self.wb)


def xml2xlsx(xml):
    parser = etree.XMLParser(target=XML2XLSXTarget(), encoding='UTF-8',
                             remove_blank_text=True)
    return etree.XML(xml, parser, )
