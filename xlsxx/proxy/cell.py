import datetime
import bisect
from typing import Generator, List

from docxx.shared import ElementProxy
from xlsxx.oxml.simpletypes import ST_CellType
from xlsxx.coord import (
    index_to_column, ref_to_coord, coord_to_ref, 
    rowref_to_index, modify_ref, split_ref
)
from xlsxx.proxy.styles import (
    NUMVAL_TYPE_INT,
    NUMVAL_TYPE_FLOAT,
    NUMVAL_TYPE_DATETIME,
    NUMVAL_TYPE_TIME
)
from xlsxx.coord import column_to_index

#
# from openpyxl
#
class EpochTime:
    def __init__(self, epoch, leap1900=False):
        self.epoch = epoch
        self.has_leap1900 = leap1900
    
    def convert_to_datetime(self, value: float):
        """
        起原経過時間から日付・時刻オブジェクトを作る。
        Params:
            value(float): 
        Returns:
            datetime.datetime
        """
        day, fraction = divmod(value, 1)
        diff = datetime.timedelta(milliseconds=round(fraction * 86400 * 1000))
        if 0 <= value < 1 and diff.days == 0:
            time = self._delta_to_time(diff.seconds, diff.microseconds)
            return datetime.datetime.combine(self.epoch, time)
        if 0 < value < 60 and self.has_leap1900:
            day += 1
        return self.epoch + datetime.timedelta(days=day) + diff

    def convert_to_time(self, value: float):
        """
        起原経過時間から時刻オブジェクトを作る。
        Params:
            value(float): 
        Returns:
            datetime.time: 日数は無視される
        """
        _day, fraction = divmod(value, 1)
        diff = datetime.timedelta(milliseconds=round(fraction * 86400 * 1000))
        return self._delta_to_time(diff.seconds, diff.microseconds)

    def _delta_to_time(self, seconds, microseconds):
        mins, seconds = divmod(seconds, 60)
        hours, mins = divmod(mins, 60)
        return datetime.time(hours, mins, seconds, microseconds)


EPOCH1899 = datetime.datetime(1899, 12, 30) 
EPOCH1904 = datetime.datetime(1904, 1, 1)   # Excel for mac: 2008 までは1904起点
WINDOWS_EXCEL_TIME = EpochTime(EPOCH1899, leap1900=True) # 1900年うるう年バグがある
OLD_MAC_EXCEL_TIME = EpochTime(EPOCH1904)
OTHER_APP_TIME = EpochTime(EPOCH1899)  # LibreOffice
TARGET_TIME = WINDOWS_EXCEL_TIME

"""
"""
class CellRow(ElementProxy):
    def __init__(self, element, sheet):
        super(CellRow, self).__init__(element, sheet)

    @property
    def sheet(self):
        return self._parent

    @property
    def workbook(self):
        return self._parent.workbook
    
    @property
    def ref(self):
        if self._element.r is None:
            raise ValueError("Bad row index")
        return str(self._element.r) # 1ベース行番号
    
    @property
    def _cells(self):
        return self._element.c_lst
    
    @property
    def cells(self):
        for elcell in self._cells:
            yield Cell(elcell, self)
        
    def get_range_cells(self, head, tail, *, emptynone=False):
        """
        範囲内のセルを取得する。要素が存在しない場合はNoneで埋める。
        Params:
            head(int): 先頭インデックス
            tail(int): 最終インデックス（含む）
            emptynone(bool): 要素が一つも無い場合Noneを返すフラグ
        Returns:
            List[Cell]:
        """
        cells = get_row_range_cell(self.ref, self._element, head, tail, emptynone=emptynone)
        if cells is None:
            return None
        return [Cell(c, self) if c is not None else None for c in cells]
        
    def cell(self, column=0):
        if column < 0:
            return None
        cells = self._cells
        elcell, inscell = find_cell_by_index(cells, column)
        if elcell is None:
            elcell, cells = self._add_cell(column, inscell, cells)
        return Cell(elcell, self)

    def new_empties_until(self, column):
        """ 指定の列に達するまでセルを末尾に追加する。
        Params:
            columnindex(int): カラム0座標
        """
        cur = len(self._cells)-1
        if cur >= column:
            return
        for i in range(column-cur):
            self._add_cell(cur+i+1, inspoint=INSERTCELL_APPEND)

    def _add_cell(self, column, inspoint=None, cells=None):
        """ 空のセルを追加する。
        Params:
            columnindex(int): カラム0座標
            inspoint(int): 挿入点
        Returns:
            Cell:
            List[Cell]:
        """
        if cells is None:
            cells = self._cells
        if inspoint is None:
            cell, inspoint = find_cell_by_index(cells, column)
        if inspoint is not None:
            cell = self._element._add_c()
            if inspoint is INSERTCELL_APPEND:
                cells.append(cell)
            else:
                inscell = cells[inspoint]
                inscell.addprevious(cell)
                cells.insert(inspoint, cell)
        cell.r = modify_ref("", row=self.ref, col=index_to_column(column))
        return cell, cells

    def write_texts(self, column_texts):
        """
        行にテキストを一度に書き込む
        Params:
            column_texts(List[Tuple[int, str]]): カラムとテキストの組のリスト
        """
        if not column_texts:
            return
        book = self.workbook
        cells = self._element.c_lst
        icell = 0
        for col, text in sorted(column_texts, key=lambda x:x[0]):
            # 書き込み先セルを順に探す
            destcell = None
            while icell < len(cells):
                cell = cells[icell]
                cellcol = ref_to_coord(cell.r)[1]
                if cellcol == col:
                    destcell = cell
                    break
                elif cellcol > col:
                    destcell, cells = self._add_cell(col, icell, cells) # cellsも更新する必要がある
                    break
                icell += 1
            else:
                destcell, cells = self._add_cell(col, INSERTCELL_APPEND, cells)
            
            # 書き込む
            set_cell_text(destcell, book, text)
    
class INSERTCELL_APPEND:
    pass

def find_cell_by_index(cells, columnindex):
    colindices = [ref_to_coord(c.r)[1] for c in cells]
    pos = bisect.bisect_left(colindices, columnindex)
    if pos < len(colindices) and colindices[pos] == columnindex:
        return cells[pos], None # 見つかった
    elif pos >= len(colindices):
        return None, INSERTCELL_APPEND # 末尾に新規追加
    else:
        return None, pos # このセルの前に新規追加

def get_row_range_cell(rowkey, element, head, tail, emptynone):
    if not isinstance(rowkey, str):
        raise ValueError("CellRow.ref has unexpected value")
    
    if tail == -1:
        celldict = {}
        maxcol = 0
        for c in element.c_lst:
            _ir, ic = ref_to_coord(c.r)
            maxcol = max(maxcol, ic)
            celldict[ic] = c
        cells = [celldict.get(i, None) for i in range(head, maxcol+1)]
    else:
        celldict = {}
        for c in element.c_lst:
            _ir, ic = ref_to_coord(c.r)
            celldict[ic] = c
        cells = [celldict.get(i, None) for i in range(head, tail+1)]
    
    if not celldict and emptynone:
        return None
    return cells

#
#
# セルの実装関数
#
#
def get_cell_value(element, book):
    """
    Params:
        element(Element): 要素
        book(Proxy): ワークブック
    """
    val = element.value()
    if val is None:
        return None
    celltype = element.t
    if celltype in (ST_CellType.SHARED_STRING): # 共有文字列型
        return get_cell_shared_text(book, val)
    elif celltype == ST_CellType.NUMBER:
        numtype = get_cell_number_value_type(element, book)
        if numtype == NUMVAL_TYPE_DATETIME: # 数値 - 日付型
            return get_cell_datetime_value(element)
        elif numtype == NUMVAL_TYPE_TIME:   # 数値 - 時刻型
            return get_cell_time_value(element)
        elif numtype == NUMVAL_TYPE_FLOAT:  # 数値 - 浮動小数点
            return float(val) 
        else:                               # 数値 - 整数
            return int(val) 
    else: # その他の型、ブール型、エラー型、空のセルはそのまま
        return val

def get_cell_text(element, book, shared_strings_map=None):
    """ 
    文字列の値を取り出す
    Params:
        element(Element): 要素 
        book(Proxy): ワークブックの要素
        *shared_strings_map(Dict[int, str]): あらかじめ取得済みの文字列マップ 
    """
    v = element.value()
    if v is None:
        return ""
    if element.t == ST_CellType.SHARED_STRING:
        return get_cell_shared_text(book, v, shared_strings_map)
    else:
        v = get_cell_value(element, book)
        if v is None:
            return ""
        return str(v)

def get_cell_shared_text(book, text, shared_strings_map=None):
    if len(text)==0:
        return ""
    if text[0] == "M":
        # 一度変更されたテキストを読み込む
        _cell, t = book.shared_strings._get_pending_text(int(text[1:]))
        return t
    else:
        index = int(text)
        if shared_strings_map:
            return shared_strings_map.get(index, "")
        else:
            # shared-stringのテーブルから読み込む   
            return book.shared_strings.get_text(index)

def get_cell_datetime_value(element, time=None):
    if time is None: time = TARGET_TIME
    fl = float(element.value())
    return time.convert_to_datetime(fl)

def get_cell_time_value(element, time=None):
    if time is None: time = TARGET_TIME
    fl = float(element.value())
    return time.convert_to_time(fl)
    
def get_cell_number_value_type(element, book):
    style_index = element.s
    return book.style_sheet.get_format(style_index).number_value_type

#
#
#
def set_cell_text(element, book, value):
    if not isinstance(value, str):
        value = str(value)
    element.t = ST_CellType.SHARED_STRING
    modid = element.get_or_add_v().text
    if modid and modid[0] == "M":
        modidnum = int(modid[1:])
        book.shared_strings._set_pending_text(modidnum, element, value)
    else:
        modid = book.shared_strings._set_pending_text(-1, element, value)
        element.v.text = "M{}".format(modid)

"""
"""
class Cell(ElementProxy):
    def __init__(self, element, row):
        super(Cell, self).__init__(element, row) # parent == CellRow
        self._xfmt = None  # スタイルバッファ

    @property
    def cellrow(self):
        return self._parent

    @property
    def workbook(self):
        return self._parent.workbook
    
    @property
    def _type(self):
        return self._element.t
    
    def is_string(self):
        return self._type in (
            ST_CellType.SHARED_STRING, ST_CellType.STR, ST_CellType.INLINE_STR
        )

    def is_number(self):
        return self._type == ST_CellType.NUMBER
    
    def is_boolean(self):
        return self._type == ST_CellType.BOOLEAN
    
    def is_error(self):
        return self._type == ST_CellType.ERROR
    
    def is_integer(self):
        return self.is_number() and self.style.number_value_type == NUMVAL_TYPE_INT
    
    def is_float(self):
        return self.is_number() and self.style.number_value_type == NUMVAL_TYPE_FLOAT
    
    def is_datetime(self):
        return self.is_number() and self.style.number_value_type == NUMVAL_TYPE_DATETIME
    
    def is_time(self):
        return self.is_number() and self.style.number_value_type == NUMVAL_TYPE_TIME
        
    def get_value(self):
        return get_cell_value(self._element, self.workbook)
    
    def clear_value(self):
        # とりあえず空文字列を代入。これでいいのか？
        self._element.t = ST_CellType.STR
        self._element.get_or_add_v().text = ""
        self._v = None
    
    def get_text(self, shared_strings_map=None):
        return get_cell_text(self._element, self.workbook, shared_strings_map)
    
    def get_datetime_value(self, time=None):
        return get_cell_datetime_value(self._element, time)

    def get_time_value(self, time=None):
        return get_cell_time_value(self._element, time)
    
    @property
    def raw(self):
        return self._element.value()

    @property
    def value(self):
        return self.get_value()
    
    @property
    def text(self):
        return self.get_text()
    
    @text.setter
    def text(self, t):
        set_cell_text(self._element, self.workbook, t)

    @property
    def empty(self):
        return self._element.v is None

    def is_empty_cell(self):
        if self.empty:
            return True
        return len(self.text.strip()) == 0

    @property
    def ref(self):
        return self._element.r

    @property
    def coord(self):
        return ref_to_coord(self.ref)

    @property
    def row(self):
        return self.coord[0]

    @property
    def row_letter(self):
        return split_ref(self.ref)[1]
    
    @property
    def column(self):
        return self.coord[1]
    
    @property
    def column_letter(self):
        return split_ref(self.ref)[0]
    
    #
    # スタイル
    #
    @property
    def style(self):
        f = self._xfmt
        if f is None:
            style_index = self._element.s
            f = self.workbook.style_sheet.get_format(style_index)
            self._xfmt = f
        return f
    


class CellRange:
    """
    """
    def __init__(self, sheet, lefttop, rightbottom):
        self._sheet = sheet
        self._lt = lefttop
        self._rb = rightbottom

    def __iter__(self):
        return self.cells()
    
    @property
    def sheet(self):
        return self._sheet
    
    @property
    def coords(self):
        return (self._lt, self._rb)

    def row_cells(self, *, sequence=False, emptynone=False) -> Generator[List[Cell], None, None]:
        """
        行ごとにセルのリストを返す。
        Yields:
            List[Cell]:
        """
        r1, c1 = self._lt
        r2, c2 = self._rb
        for row in self._sheet.get_range_rows(r1, r2):
            cells = row.get_range_cells(c1, c2, emptynone=emptynone)
            if sequence and cells is None:
                break
            yield cells
    
    def column_cells(self):
        """
        列ごとにセルのリストを返す。
        Yields:
            List[Cell]
        """
        r1, c1 = self._lt
        r2, c2 = self._rb
        col = [None for _ in range(r2-r1+1)]
        cols = [col[:] for _ in range(c2-c1+1)]
        for j, row in enumerate(self._sheet.get_range_rows(r1, r2)):
            rowcells = row.get_range_cells(c1, c2)
            for i, cell in enumerate(rowcells):
                cols[i][j] = cell
        
        for col in cols:
            yield col
    
    def cells(self, *, sequence=False, emptynone=False):
        """
        すべてのセルを順に返す。
        """
        for cells in self.row_cells(emptynone=emptynone):
            for cell in cells:
                yield cell

#
#
#
def get_range_text(sheet, lefttop, rightbottom, *, iterbreak=True, strmap=None):
    """
    Params:
        sheet(Proxy): ワークシート
    Returns:    
        List[Tuple[Str, Tuple[int, int]]]: (テキスト、座標)のリスト　行の区別のない一次元のリスト
    """
    texts = []
    r1, c1 = lefttop
    r2, c2 = rightbottom
    book = sheet.workbook
    rlast = r2 + 1 if r2 >= 0 else None
    for row in sheet.element.sheetData.row_lst[r1:rlast]:
        rowletter = str(row.r)
        cs = get_row_range_cell(rowletter, row, c1, c2, iterbreak)
        #if iterbreak and cs is None:
        #    break
        if cs is None:
            continue
        for ci, cell in enumerate(cs, start=c1):
            if cell is not None:
                text = get_cell_text(cell, book, strmap)
            else:
                text = ""
            coord = (rowref_to_index(rowletter), ci)
            texts.append((text, coord))
    return texts
