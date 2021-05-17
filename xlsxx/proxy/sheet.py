# encoding: utf-8

"""
|Document| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docxx.shared import ElementProxy, AttributeProperty
from xlsxx.coord import ref_to_coord, coord_to_ref, range_ref_to_coord, column_to_index, split_ref, get_range_coord
from xlsxx.proxy.cell import Cell, CellRow, CellRange

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
        self._entry = entry

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
    def rId(self):
        return self._entry.rel_id
    
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
    def last_row(self):
        """
        最も後ろにあるセルの行番号を返す
        Returns:
            int: 0ベースの行番号
        """
        rowlst = self._element.sheetdata.row_lst
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
        rowlst = self._element.sheetdata.row_lst
        if not rowlst:
            return 0
        maxlen = max(len(x.cell_lst) for x in rowlst)
        return maxlen-1 if maxlen>0 else 0
    
    @property
    def columns(self):
        """
        列を表すオブジェクトをすべて取得する。
        Returns:
            List[Column]
        """
        cols = self._element.cols_lst
        if not cols:
            raise ValueError("no columns")
        return [Column(x, self) for x in cols[0].col_lst]
    
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
        cols = self._element.cols_lst[column]
        if not cols:
            raise ValueError("no columns")
        return Column(cols[0].col_lst[column], self)
    
    @property
    def rows(self):
        """
        行を表すオブジェクトをすべて取得する。
        Params:
            row(int): 行の番号
        Returns:
            List[CellRow]:
        """
        return [CellRow(el, self._workbook) for el in self._element.sheetdata.row_lst]
    
    def get_range_rows(self, head, tail):
        """
        範囲内の行オブジェクトを取得する。
        Params:
            head(int): 開始インデックス
            tail(int): 最終インデックス（含む）
        Returns:
            List[CellRow]:
        """
        return [CellRow(el, self._workbook) for el in self._element.sheetdata.row_lst[head:tail+1]]

    def row(self, row):
        """
        行を表すオブジェクトを取得する。
        Params:
            row(int): 行の番号
        Returns:
            CellRow:
        """
        elrow = self._element.sheetdata.row_lst[row]
        return CellRow(elrow, self._workbook)
    
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

    def range(self, lefttop, rightbottom=None, *, rownum=None, columnnum=None, orientation=None):
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
        return CellRange(self, p1, p2, orientation=orientation)

    def vertical_range(self, lefttop, length=None):
        """
        開始点から縦1列の範囲を作成する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            CellRange:
        """
        if length is None:
            tail = self.last_row
        else:
            if length <= 0:
                raise ValueError("長さは1以上必要です")
            tail = lefttop[0] + length - 1
        return self.range(lefttop, (tail, lefttop[1]), orientation="v")
        
    def horizontal_range(self, lefttop, length=None):
        """
        開始点から横1行の範囲を作成する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            CellRange
        """
        if length is None:
            tail = self.last_column
        else:
            if length <= 0:
                raise ValueError("長さは1以上必要です")
            tail = lefttop[1] + length - 1
        return self.range(lefttop, (lefttop[0], tail), orientation="h")    
    
    def allocate_range(self, lefttop, rightbottom=None):
        """
        範囲内に空のセルを追加する。
        Params:
            lefttop(Tuple/str): 開始点／範囲参照
            rightbottom(Tuple/str): 終了点（境界を含む）
        """
        p1, p2 = get_range_coord(lefttop, rightbottom)
        rmin, _cmin = p1
        rmax, cmax = p2
        curmax = len(self._element.sheetdata.row_lst)-1 # 空の場合は-1になる
        # 空の行を追加する
        for ri in range(rmax-curmax):
            self.add_row(ri)
        # 各行を列方向に進捗する
        for ri in range(rmin, rmax+1):
            self.row(ri).new_empties_until(cmax)

    def add_row(self, index):
        """
        空の行を追加する。
        """
        row = self._element.sheetdata._add_row()
        row.ref = index + 1

"""
"""
class Column(ElementProxy):
    min = AttributeProperty("min")
    max = AttributeProperty("max")
    width = AttributeProperty("width")
    