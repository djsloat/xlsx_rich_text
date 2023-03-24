from functools import cached_property

import lxml.etree
from xlsx_rich_text.cell.datacell import DataCell
from xlsx_rich_text.ooxml_ns import ns


class Record:
    def __init__(self, element: lxml.etree._Element, sheet):
        self.element = element
        self._sheet = sheet
        self.row_num = self.element.xpath("string(@r)", **ns)
        self.record_num = str(int(self.row_num) - int(self._sheet.header_row) - 1)

    def __repr__(self):
        return f"{self.__class__.__name__}(row={self.row_num})"

    def __getitem__(self, key):
        return self.col.get(key)

    def __iter__(self):
        return iter(self.col)

    def __len__(self):
        return len(self.col)

    def __lt__(self, other):
        return self.record_num < other.record_num

    @cached_property
    def col(self):
        return {
            cell.column: cell
            for cell in [DataCell(c, self._sheet) for c in self.element]
        }
