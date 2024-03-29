# encoding: utf-8

"""
|Document| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)
from collections import defaultdict

from docxx.shared import ElementProxy, AttributeProperty
from docxx.element import remove_element, query
from xlsxx.coord import (
    ref_to_coord, coord_to_ref, range_ref_to_coord, column_to_index, 
    get_coord, get_range_coord, modify_range_ref, rowref_to_index, index_to_rowref
)
from xlsxx.proxy.cell import Cell, CellRow, CellRange, get_range_values

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
        return [CellRow(el, self) for el in self._rows]
    
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
        return [CellRow(el, self) for el in self._rows[head:last]]
    
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
        return CellRow(self._rows[row], self)
    
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

    def range(self, lefttop, rightbottom=None, *, rownum=None, columnnum=None):
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
        return CellRange(self, p1, p2)
    
    def get_range_text(self, lefttop, rightbottom=None, *, rownum=None, columnnum=None, readingcells=None):
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
        return get_range_values(self, p1, p2, readingcells or ReadingCells())

    def vertical_range(self, lefttop, length=None):
        """
        開始点から縦1列の範囲を作成する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            CellRange:
        """
        r1, c1 = get_coord(lefttop)
        tail, stop = _vertical_tail(self, (r1, c1), length)
        return self.range((r1, c1), (tail, c1))
    
    def get_vertical_range_text(self, lefttop, length=None, *, readingcells=None):
        """
        開始点から縦1列の範囲のテキストを取得する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            List[Tuple[Str, Tuple[int, int]]]:
        """
        r1, c1 = get_coord(lefttop)
        tail, stop = _vertical_tail(self, (r1, c1), length)
        return get_range_values(self, (r1, c1), (tail, c1), readingcells or ReadingCells())
        
    def horizontal_range(self, lefttop, length=None):
        """
        開始点から横1行の範囲を作成する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            CellRange:
        """
        r1, c1 = get_coord(lefttop)
        tail = _horizontal_tail(self, (r1, c1), length)
        return self.range((r1, c1), (r1, tail))
    
    def get_horizontal_range_text(self, lefttop, length=None, *, readingcells=None):
        """
        開始点から横1行の範囲のテキストを取得する。
        Params:
            lefttop(Tuple/str): 開始点のセル
            length(int): 増分
        Returns:
            List[Tuple[Str, Tuple[int, int]]]:
        """
        r1, c1 = get_coord(lefttop)
        tail, stop = _horizontal_tail(self, (r1, c1), length)
        return get_range_values(self, (r1, c1), (r1, tail), readingcells or ReadingCells())    
    
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
        # 空の行を追加する
        curmax = len(self._rows)-1 # 空の場合は-1になる
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

    def insert_row(self, posrow, index, *, prev=False):
        """
        指定要素の後ろに行を新規作成する。
        Params:
            index(int): 0ベース列番号
        """
        rownew = self._element.sheetData._new_row()
        rownew.r = index + 1
        if prev:
            posrow.addprevious(rownew)
        else:
            posrow.addnext(rownew)
        return rownew

    def allocate_rows(self, rhead, rtail):
        """
        空の行を後ろに複数追加する。
        途中に空行があれば埋める。
        Params:
            rhead(int): 0ベース行番号
            rtail(int): 0ベース行番号
        """
        if self._rows:
            elindex = 0
            nextirow = rhead
            while elindex < len(self._rows):
                elrow = self._rows[elindex]
                irow = rowref_to_index(elrow.r) # 0ベース行番号に変換
                if irow > rtail:
                    # 後ろに要素の空きがある場合に対応
                    for di in range(nextirow, rtail+1): # rtailにも要素が存在しないので含める
                        self.insert_row(elrow, di, prev=True)
                    elindex += (rtail-nextirow+1) + 1
                    nextirow = rtail + 1
                    break
                elif rhead <= irow:
                    if irow > nextirow:
                        # 飛んでいる - 差分を全て要素として追加する
                        for di in range(nextirow, irow): # irowには要素が存在するので含めない
                            self.insert_row(elrow, di, prev=True)
                        elindex += (irow-nextirow) + 1
                        nextirow = irow + 1
                        continue
                elindex += 1
            
            # 残り
            elrow = self._rows[-1]
            irow = rowref_to_index(elrow.r)
            if irow < rtail:
                for di in range(irow+1, rtail+1):
                    elrow = self.insert_row(elrow, di)
        
        else:
            for i in range(rhead, rtail+1):
                self.add_row(i)

    #
    #
    #
    def write_rows_text(self, lefttop, rows):
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

    def write_cells(self, writings):
        """
        座標とテキストの組のリストから一気にセルへの書き込みを行う。
        Params:
            writings(WritingCells): 座標とテキストの組のリスト
        """
        if not isinstance(writings, WritingCells):
            raise TypeError("'writing' must be WritingCells instance")
        if writings.empty():
            return
        
        # 必要な行を確保する
        mirow, _icol = writings.mincoord()
        mxrow, _mcol = writings.maxcoord()
        self.allocate_rows(mirow, mxrow) 

        cellrows = {}
        for cellrow in self._rows:
            rowindex = int(cellrow.r)-1
            if writings.has_row(rowindex):
                cellrows[rowindex] = CellRow(cellrow, self)

        for row, coltexts in writings.rows():
            cellrow = cellrows.get(row)
            if cellrow is None:
                raise ValueError("Row '{}' is not allocated".format(index_to_rowref(row)))
            cellrow.write(coltexts, as_values=writings.as_values())

    def writing_cells(self, *, as_values=False):
        return WritingCells(as_values=as_values)

    def reading_cells(self, *, sequential=False, nofetch=False, as_values=False):
        return ReadingCells(sequential=sequential, nofetch=nofetch, as_values=as_values)
    

"""
"""
class Column(ElementProxy):
    min = AttributeProperty("min")
    max = AttributeProperty("max")
    width = AttributeProperty("width")

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
    else:
        if length <= 0:
            raise ValueError("長さは1以上必要です")
        tail = lefttop[1] + length - 1
    return tail


#
#
#
class ComlumnDef:
    def __init__(self, ref, type):
        self.ref = ref
        self.type = type
    
#
#
#
class WritingCells:
    def __init__(self, *, as_values=False):
        self._rows = defaultdict(list)
        self._mincoord = None
        self._maxcoord = None
        self._asvalues = as_values

    def empty(self):
        return len(self._rows) == 0

    def add(self, coord, value):
        if not isinstance(coord, tuple):
            raise TypeError("coord must be tuple")
        row, col = coord
        self._rows[row].append((col, value))
        self._mincoord = (row, col) if self._mincoord is None else min(self._mincoord, (row, col)) 
        self._maxcoord = (row, col) if self._maxcoord is None else max(self._maxcoord, (row, col))

    def mincoord(self):
        return self._mincoord
    
    def maxcoord(self):
        return self._maxcoord
    
    def rows(self):
        return self._rows.items()

    def has_row(self, index):
        return index  in self._rows

    def as_values(self):
        return self._asvalues


class ReadingCells:
    def __init__(self, *, 
        sequential=False, 
        nofetch=False, 
        as_values=False,
        column_types=None
    ):
        self.sequential = sequential
        self.nofetch = nofetch
        self.asvalues = as_values
        self._strmap = None
        self._coltypes = column_types

    def set_strmap(self, strmap):
        self._strmap = strmap

    def set_column_type(self, columnref, type):
        # フォーマットを指定する
        if self._coltypes is None:
            self._coltypes = {}        
        from xlsxx.proxy.styles import (
            NUMVAL_TYPE_DATETIME, NUMVAL_TYPE_TIME, NUMVAL_TYPE_INT, NUMVAL_TYPE_FLOAT
        )
        t = {
            "i" : NUMVAL_TYPE_INT,
            "f" : NUMVAL_TYPE_FLOAT,
            "d" : NUMVAL_TYPE_DATETIME,
            "t" : NUMVAL_TYPE_TIME,
            "_" : NUMVAL_TYPE_INT,
            "-" : NUMVAL_TYPE_INT,
        }.get(type.lower())
        if t is None:
            raise ValueError('[I]nt/[F]loat/[D]atetime/[T]ime/_のいずれかを指定してください')
        self._coltypes[columnref] = t

    @property
    def strmap(self):
        return self._strmap

    @property
    def default_value(self):
        if self.asvalues:
            return None
        else:
            return ""
    
    @property
    def column_number_types(self):
        return self._coltypes


#
# 即座に値を読み書きする関数
#
def read_sheet_rows(sheet, start, tailcolumn, tailrow=-1, *, readingcells:ReadingCells=None):
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
    if not readingcells.nofetch:
        readingcells.set_strmap(sheet.workbook.part.fetch_textmap())
    
    from xlsxx.coord import get_range_coord
    (startrow, startcolumn), (tailrow, tailcolumn) = get_range_coord(start, (tailrow, tailcolumn))
    values = get_range_values(sheet, (startrow, startcolumn), (tailrow, tailcolumn), readingcells or ReadingCells())
    rows = []
    i = 0
    #
    sequential = readingcells.sequential
    default_value = readingcells.default_value
    while i < len(values):
        row = [default_value for _ in range(tailcolumn-startcolumn+1)]
        for j in range(len(row)):
            value, _ref = values[i+j]
            row[j] = value
        if sequential and all(x == default_value for x in row):
            # 空の行が来たら読み込みを中止する
            break
        i += (j + 1)
        rows.append(row)
    return rows

