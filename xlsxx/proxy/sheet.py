# encoding: utf-8

"""
|Document| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)
from typing import Iterable

from docxx.shared import ElementProxy, AttributeProperty
from docxx.element import remove_element, query
from xlsxx.coord import (
    ref_to_coord, coord_to_ref, range_ref_to_coord, column_to_index, 
    get_coord, get_range_coord, modify_range_ref
)
from xlsxx.proxy.cell import Cell, CellRow, CellRange, get_range_text

"""
"""
class Worksheet(ElementProxy):
    """
    Sequence [1..1]
        ssml:sheetPr [0..1]    Worksheet Properties
        ssml:dimension [0..1]    Worksheet Dimensions
        ssml:sheetViews [0..1]    Sheet Views
        ssml:sheetFormatPr [0..1]    Sheet Format Properties
        ssml:cols [0..*]    Column Information
        ssml:sheetData [1..1]    Sheet Data
        ssml:sheetCalcPr [0..1]    Sheet Calculation Properties
        ssml:sheetProtection [0..1]    Sheet Protection
        ssml:protectedRanges [0..1]    Protected Ranges
        ssml:scenarios [0..1]    Scenarios
        ssml:autoFilter [0..1]    AutoFilter
        ssml:sortState [0..1]    Sort State
        ssml:dataConsolidate [0..1]    Data Consolidate
        ssml:customSheetViews [0..1]    Custom Sheet Views
        ssml:mergeCells [0..1]    Merge Cells
        ssml:phoneticPr [0..1]    Phonetic Properties
        ssml:conditionalFormatting [0..*]    Conditional Formatting
        ssml:dataValidations [0..1]    Data Validations
        ssml:hyperlinks [0..1]    Hyperlinks
        ssml:printOptions [0..1]    Print Options
        ssml:pageMargins [0..1]    Page Margins
        ssml:pageSetup [0..1]    Page Setup Settings
        ssml:headerFooter [0..1]    Header Footer Settings
        ssml:rowBreaks [0..1]    Horizontal Page Breaks
        ssml:colBreaks [0..1]    Vertical Page Breaks
        ssml:customProperties [0..1]    Custom Properties
        ssml:cellWatches [0..1]    Cell Watch Items
        ssml:ignoredErrors [0..1]    Ignored Errors
        ssml:smartTags [0..1]    Smart Tags
        ssml:drawing [0..1]    Drawing
        ssml:legacyDrawing [0..1]    Legacy Drawing
        ssml:legacyDrawingHF [0..1]    Legacy Drawing Header Footer
        ssml:picture [0..1]    Background Image
        ssml:oleObjects [0..1]    OLE Objects
        ssml:controls [0..1]    Embedded Controls
        ssml:webPublishItems [0..1]    Web Publishing Items
        ssml:tableParts [0..1]    Table Parts
        ssml:extLst [0..1]    Future Feature Storage Area
    """

    __slots__ = ('_part', '_workbook', '_entry')

    def __init__(self, element, part, workbook, entry):
        super(Worksheet, self).__init__(element)
        self._part = part
        self._workbook = workbook
        self._entry = entry # CT_Sheet

    @property
    def workbook(self):
        """
        このシートを含むワークブックを得る。
        """
        return self._workbook

    @property
    def part(self):
        """
        The |DocumentPart| object of this document.
        """
        return self._part
 
    @property
    def name(self):
        return self._entry.name
    
    @name.setter
    def name(self, n):
        self._entry.name = n
    
    @property
    def id(self):
        return self._entry.sheetId
    
    @property
    def relid(self):
        return self._entry.rid
    
    @property
    def dimension(self):
        dim = self._element.dimension
        if dim is not None:
            _beg, end = range_ref_to_coord(dim.ref)
            return end
    
    @property
    def limit_row(self):
        """
        現在の行番号の限界値を返す
        Returns:
            int: 0ベースの行番号
        """
        dim = self.dimension
        if dim is None:
            return None
        return dim[0]
    
    @property
    def limit_column(self):
        """
        現在の列番号の限界値を返す
        Returns:
            int: 0ベースの列番号
        """
        dim = self.dimension
        if dim is None:
            return None
        return dim[1]
    
    @property
    def _rows(self):
        return self._element.sheetData.row_lst
    
    @property
    def last_row(self):
        """
        最も後ろにあるセルの行番号を返す
        Returns:
            int: 0ベースの行番号
        """
        rowlst = self._rows
        if not rowlst:
            return 0
        return len(rowlst)-1
    
    @property
    def last_column(self):
        """
        最も後ろにあるセルの列番号を返す
        Returns:
            int: 0ベースの列番号
        """
        rowlst = self._rows
        if not rowlst:
            return 0
        maxlen = max(len(x.c_lst) for x in rowlst)
        return maxlen-1 if maxlen>0 else 0
    
    @property
    def _columns(self):
        cols = self._element.cols_lst
        if not cols:
            return []
        return cols[0].col_lst # colsの全要素を結合する？
    
    @property
    def columns(self):
        """
        列を表すオブジェクトをすべて取得する。
        Returns:
            List[Column]
        """
        return [Column(x, self) for x in self._columns]
    
    def column(self, column):
        """
        列を表すオブジェクトを取得する。
        Params:
            column(int/str): 列の番号／文字
        Returns:
            Column:
        """
        if isinstance(column, str):
            column = column_to_index(column)
        cols = self._columns
        if column < 0 or len(cols) <= column:
            raise IndexError("column")
        return Column(cols[column], self)
    
    @property
    def rows(self):
        """
        行を表すオブジェクトをすべて取得する。
        Params:
            row(int): 行の番号
        Returns:
            List[CellRow]:
        """
        return [CellRow(el, self._workbook) for el in self._rows]
    
    def get_range_rows(self, head, tail):
        """
        範囲内の行オブジェクトを取得する。
        Params:
            head(int): 開始インデックス
            tail(int/None): 最終インデックス（含む）／最後まで
        Returns:
            List[CellRow]:
        """
        if tail is None:
            last = None
        else:
            last = tail + 1
        return [CellRow(el, self._workbook) for el in self._rows[head:last]]
    
    def row(self, row):
        """
        行を表すオブジェクトを取得する。
        Params:
            row(int): 行の番号
        Returns:
            CellRow:
        """
        if row < 0 or len(self._rows) <= row:
            return None
        return CellRow(self._rows[row], self._workbook)
    
    def minimize_dimension(self):
        # セルと行の数を調べ、最小限の寸法を決める
        lastrowref = None
        lastrowoffset = 0
        for row in self.rows:
            empty = True
            for cell in row.cells:
                if not cell.empty: # 空文字列であっても要素が存在する場合は偽とする
                    empty = False
                    break
            if empty:
                break
            lastrowref = row.ref
            lastrowoffset += 1

        if lastrowref is None:
            return
        
        # ディメンジョンの修正
        dim = self._element.dimension
        if dim is not None:
            dim.ref = modify_range_ref(dim.ref, tailrow=lastrowref)
        
        # 空の行を削除
        for elem in self._rows[lastrowoffset:]:
            remove_element(self._element.sheetData, query(elem))

    
    #
    # セルの取得
    #
    def cell(self, row, column=None):
        """
        セルをひとつ取得する。
        Params:
            row(int/str): 行の番号／セル参照
            column(int): 列の番号
        Returns:
            Cell:
        """
        if column is None:
            if not isinstance(row, str):
                raise TypeError("row")
            r, c = ref_to_coord(row)
            return self.row(r).cell(c)
        else:
            return self.row(row).cell(column)

    def range(self, lefttop, rightbottom=None, *, rownum=None, columnnum=None, stop=True):
        """
        矩形のセル範囲を作成する。
        Params:
            lefttop(Tuple/str): 矩形範囲の左上のセル
            rightbottom(Tuple/str): 矩形範囲の右下のセル（境界を含む）
            rownum(int): 左上からの行の増分
            columnnum(int): 左上からの列の増分
        Returns:
            CellRange:
        """
        p1, p2 = get_range_coord(lefttop, rightbottom, rownum=rownum, columnnum=columnnum)
        return CellRange(self, p1, p2, iterbreak=stop)
    
    def get_range_text(self, lefttop, rightbottom=None, *, rownum=None, columnnum=None, stop=True, strmap=None):
        """
        矩形のセル範囲のテキストを取得する。
        Params:
            lefttop(Tuple/str): 矩形範囲の左上のセル
            rightbottom(Tuple/str): 矩形範囲の右下のセル（境界を含む）
            rownum(int): 左上からの行の増分
            columnnum(int): 左上からの列の増分
        Returns:
            List[Tuple[Str, Str]]: セル文字列、セル参照のタプルのリスト
        """
        p1, p2 = get_range_coord(lefttop, rightbottom, rownum=rownum, columnnum=columnnum)
        return get_range_text(self, p1, p2, iterbreak=stop, strmap=strmap)

    def vertical_range(self, lefttop, length=None):
        """
        開始点から縦1列の範囲を作成する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            CellRange:
        """
        tail, stop = _vertical_tail(self, lefttop, length)
        return self.range(lefttop, (tail, lefttop[1]), stop=stop)
    
    def get_vertical_range_text(self, lefttop, length=None, *, strmap=None):
        """
        開始点から縦1列の範囲のテキストを取得する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            List[Str]:
        """
        tail, stop = _vertical_tail(self, lefttop, length)
        return self.get_range_text(lefttop, (tail, lefttop[1]), stop=stop, strmap=strmap)
        
    def horizontal_range(self, lefttop, length=None):
        """
        開始点から横1行の範囲を作成する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            CellRange:
        """
        tail, stop = _horizontal_tail(self, lefttop, length)
        return self.range(lefttop, (lefttop[0], tail), stop=stop)    
    
    def get_horizontal_range_text(self, lefttop, length=None, *, strmap=None):
        """
        開始点から横1行の範囲のテキストを取得する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            List[Str]:
        """
        tail, stop = _horizontal_tail(self, lefttop, length)
        return self.get_range_text(lefttop, (lefttop[0], tail), stop=stop, strmap=strmap)    
    
    #
    # 書き込み
    #
    def allocate_range(self, lefttop, rightbottom=None):
        """
        範囲内に空のセルを追加する。
        Params:
            lefttop(Tuple/str): 開始点／範囲参照
            rightbottom(Tuple/str): 終了点（境界を含む）
        Returns:
            List[CellRow]:
        """
        p1, p2 = get_range_coord(lefttop, rightbottom)
        rmin, _cmin = p1
        rmax, cmax = p2
        curmax = len(self._rows)-1 # 空の場合は-1になる
        # 空の行を追加する
        for i in range(rmax-curmax):
            self.add_row(curmax+i+1)
        # 各行を列方向に進捗する
        rows = []
        for i in range(rmin, rmax+1):
            row = self.row(i)
            row.new_empties_until(cmax)
            rows.append(row)
        return rows

    def add_row(self, index):
        """
        空の行を追加する。
        Params:
            index(int): 0ベース列番号
        """
        row = self._element.sheetData._add_row()
        row.r = index + 1

    def write_rows(self, lefttop, rows):
        """
        行の値のリストを流し込む。
        Params:
            lefttop(Tuple/str): 開始点／範囲参照
            rows(Tuple[Tuple[str]]): 行のリスト、長さはそろわなくてよい
        """
        maxrowlen = max(len(x) for x in rows)
        r1, c1 = get_coord(lefttop)
        r2 = r1 + len(rows)
        c2 = c1 + maxrowlen
        cellrows = self.allocate_range((r1, c1), (r2, c2))
        for row, cellrow in zip(rows, cellrows):
            for ci, cell in enumerate(cellrow.cells):
                if ci < len(row):
                    cell.text = row[ci]

#
#
#
def _vertical_tail(proxy, lefttop, length):
    if length is None:
        tail = proxy.last_row
        stop = True
    else:
        if length <= 0:
            raise ValueError("長さは1以上必要です")
        tail = lefttop[0] + length - 1
        stop = False
    return tail, stop

def _horizontal_tail(proxy, lefttop, length):
    if length is None:
        tail = proxy.last_column
        stop = True
    else:
        if length <= 0:
            raise ValueError("長さは1以上必要です")
        tail = lefttop[1] + length - 1
        stop = False
    return tail, stop


"""
"""
class Column(ElementProxy):
    min = AttributeProperty("min")
    max = AttributeProperty("max")
    width = AttributeProperty("width")
    

#
# 即座に値を読み書きする関数
#
def read_sheet_vertical(sheet, start, tailcolumn, tailrow=-1, *, sequence=False, nofetch=False):
    """
    シートから縦方向に値を読みだす
    Params:
        sheet(xlsxx.proxy.Worksheet): 対象シート
        start(Tuple[str|int, str|int]|str): 開始位置の座標
        tailcolumn(str|int): 終了位置の列参照・列番号（含む）
        tailrow(str|int) : 終了位置の列参照・列番号（含む）
    Returns:
        List[List[Str]]: 文字列型の値が列の数だけ入った行のリスト
    """
    if nofetch:
        strmap = None
    else:
        strmap = sheet.workbook.part.fetch_textmap()
    
    from xlsxx.coord import get_range_coord
    (startrow, startcolumn), (tailrow, tailcolumn) = get_range_coord(start, (tailrow, tailcolumn))
    texts = sheet.get_range_text((startrow, startcolumn), (tailrow, tailcolumn), strmap=strmap)
    rows = []
    i = 0
    while i < len(texts):
        row = ["" for _ in range(tailcolumn-startcolumn+1)]
        for j in range(len(row)):
            value, _ref = texts[i+j]
            row[j] = value
        if sequence and all(len(x) == 0 for x in row):
            break
        i += (j + 1)
        rows.append(row)
    return rows

