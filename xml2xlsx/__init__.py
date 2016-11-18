# -*- coding: utf-8 -*-
import logging

from lxml import etree
from datetime import datetime

from decimal import Decimal, InvalidOperation

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.writer.dump_worksheet import WriteOnlyCell
from openpyxl.writer.excel import save_virtual_workbook

logger = logging.getLogger(__name__)

LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def excel_style(row, col):
    """ Convert given row and column number to an Excel-style cell name. """
    result = []
    while col:
        col, rem = divmod(col - 1, 26)
        result[:0] = LETTERS[rem]
    return ''.join(result) + str(row)


class XML2XLSXTarget(object):
    def __init__(self):
        self.wb = Workbook()
        self._current_ws = None
        self._row_buf = []
        self._cell = None
        self._cell_type = None
        self._cell_date_format = None
        self._row = 0
        self._col = 0
        self._refs = {
            'row': 1,
            'col': 1,
        }

    def start(self, tag, attrib):
        if tag == 'sheet':
            if not self._current_ws:
                self._current_ws = self.wb.active
            else:
                self._current_ws = self.wb.create_sheet()
            if 'name' in attrib:
                self._current_ws.title = attrib['name']
            self._row = 0
        elif tag == 'row':
            self._row_buf = []
            self._col = 0
        elif tag == 'cell':
            self._cell = WriteOnlyCell(self._current_ws)
            for attr, value in attrib.iteritems():
                if attr == 'font':
                    params = dict([v.split(':') for v in
                                   value.split(';') if v.strip()])
                    result = {}
                    for param, value in params.iteritems():
                        param = param.strip()
                        value = value.strip()
                        if value in ['True', 'False']:
                            result[param] = bool(value)
                        else:
                            try:
                                result[param] = int(value)
                            except:
                                result[param] = float(value)
                    font = Font(**result)
                    self._cell.font = font
                elif attr == 'ref-id':
                    self._refs[value] = excel_style(self._row + 1,
                                                    self._col + 1)
                elif attr == 'ref-append':
                    self._refs[value] = self._refs.get(value, [])
                    self._refs[value].append(excel_style(self._row + 1,
                                                         self._col + 1))
                    # self._refs[value].append()
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
            self._row += 1
            self._refs['row'] = self._row + 1
        elif tag == 'cell':
            if self._cell.value:
                stringified = {
                    k: ", ".join(v) if hasattr(v, '__iter__') else v
                    for k, v in self._refs.iteritems()
                }
                self._cell.value = self._cell.value.format(**stringified)
            if self._cell_type == 'number':
                if self._cell.value:
                    try:
                        self._cell.value = Decimal(self._cell.value)
                    except InvalidOperation:
                        pass
            elif self._cell_type == 'date':
                self._cell.value = datetime.strptime(
                        self._cell.value, self._cell_date_format).date()
            self._row_buf.append(self._cell)
            self._cell = None
            self._col += 1
            self._refs['col'] = self._col + 1

    def close(self):
        return save_virtual_workbook(self.wb)


def xml2xlsx(xml):
    parser = etree.XMLParser(target=XML2XLSXTarget(), encoding='UTF-8',
                             remove_blank_text=True)
    return etree.XML(xml, parser, )

__all__ = ['xml2xlsx']