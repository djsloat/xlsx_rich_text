"""Sheet"""

from typing import TYPE_CHECKING

from lxml.etree import _ElementTree

from xlsx_rich_text.cell.cell import Cell
from xlsx_rich_text.ooxml_ns import ns

if TYPE_CHECKING:
    from xlsx_rich_text.sheets.sheets import Sheets


class Sheet:
    """Representation of sheet.xml"""

    def __init__(self, _name: str, sheets: "Sheets") -> None:
        self._name = _name
        self._parent = sheets

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(name='{self._name}')"

    @property
    def _id(self) -> str:
        xpath = "string(w:sheets/w:sheet[@name=$_name]/@sheetId)"
        return self._parent._book_xml.xpath(xpath, _name=self._name, **ns)

    @property
    def xml(self) -> _ElementTree:
        rid_xpath = "string(w:sheets/w:sheet[@name=$_name]/@r:id)"
        rid = self._parent._book_xml.xpath(rid_xpath, _name=self._name, **ns)
        filename_xpath = "string(r1:Relationship[@Id=$_rid]/@Target)"
        filename = "xl/" + self._parent._book_rels.xpath(filename_xpath, _rid=rid, **ns)
        return self._parent._book.xml[filename]

    def cell(self, row: int | str, col: int | str) -> Cell | None:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c[@r=$_c]"
        if len(cell := self.xml.xpath(xpath, _r=row, _c=col, **ns)):
            return Cell(cell[0], self)
        return None

    def row(self, row: int | str) -> list[Cell]:
        xpath = "w:sheetData/w:row[@r=$_r]/w:c"
        return [Cell(el, self) for el in self.xml.xpath(xpath, _r=row, **ns)]

    def col(self, col: int | str) -> list[Cell]:
        xpath = "w:sheetData/w:row/w:c[@r=$_c]"
        return [Cell(el, self) for el in self.xml.xpath(xpath, _c=col, **ns)]
