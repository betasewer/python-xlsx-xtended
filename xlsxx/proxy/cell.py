import datetime

from docxx.shared import ElementProxy
from xlsxx.oxml.simpletypes import ST_CellType
from xlsxx.coord import index_to_column, ref_to_coord, coord_to_ref, modify_ref, split_ref
from xlsxx.proxy.styles import (
    NUMVAL_TYPE_NUMBER,
    NUMVAL_TYPE_DATETIME,
    NUMVAL_TYPE_TIME
)

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
    def __init__(self, element, workbook):
        super(CellRow, self).__init__(element)
        self._workbook = workbook
    
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
            yield Cell(elcell, self._element, self._workbook)
        
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
        return [Cell(c, self._element, self._workbook) if c is not None else None for c in cells]
        
    def cell(self, column=0):
        if column < 0 or len(self._cells) <= column:
            return None
        elcell = self._cells[column]
        return Cell(elcell, self._element, self._workbook)

    def new_empties_until(self, column):
        """ 空のセルを複数追加する。
        Params:
            columnindex(int): カラム0座標
        """
        cur = len(self._cells)-1
        if cur >= column:
            return
        for i in range(column-cur):
            self.add_cell(cur+i+1)

    def add_cell(self, column):
        """ 空のセルを追加する。
        Params:
            columnindex(int): カラム0座標
        """
        cell = self._element._add_c()
        cell.r = modify_ref("", row=self.ref, col=index_to_column(column))

#
def get_row_range_cell(rowkey, element, head, tail, emptynone):
    if not isinstance(rowkey, str):
        raise ValueError("CellRow.ref has unexpected value")
    if tail == -1:
        raise ValueError("Specify the column tail with an existing index value, not -1")
    celldict = {modify_ref("", rowkey, icol):i for i,icol in enumerate(range(head, tail+1))}
    cells = [None for _ in range(len(celldict))]
    empty = True
    for c in element.c_lst:
        cellref = c.r
        if cellref in celldict:
            cells[celldict[cellref]] = c
            empty = False
    if empty and emptynone:
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
    if element.v is None:
        return None
    celltype = element.t
    if celltype in (ST_CellType.SHARED_STRING): # 共有文字列型
        return get_cell_shared_text(element.v, book)
    elif celltype == ST_CellType.NUMBER:
        if get_cell_number_format(element, book).type == NUMVAL_TYPE_DATETIME: # 数値 - 日付型
            return get_cell_datetime_value(element)
        elif get_cell_number_format(element, book).type == NUMVAL_TYPE_TIME:   # 数値 - 時刻型
            return get_cell_time_value(element)
        else:                                               # 数値 - その他の型
            v = element.v
            return float(v.text)
    else: # その他の型、ブール型、エラー型、空のセル
        v = element.v
        if v is None:
            return None
        return v.text

def get_cell_text(element, book, shared_strings_map=None):
    """ 
    文字列の値を取り出す
    Params:
        element(Element): 要素 
        book(Proxy): ワークブックの要素
        *shared_strings_map(Dict[int, str]): あらかじめ取得済みの文字列マップ 
    """
    v = element.v
    if v is None:
        return ""
    if element.t == ST_CellType.SHARED_STRING:
        return get_cell_shared_text(v, book, shared_strings_map)
    else:
        v = get_cell_value(element, book)
        if v is None:
            return ""
        return str(v)

def get_cell_shared_text(value, book, shared_strings_map=None):
    if len(value.text)==0:
        return ""
    if value.text[0] == "M":
        # 一度変更されたテキストを読み込む
        _cell, t = book.shared_strings._get_pending_text(int(value.text[1:]))
        return t
    else:
        index = int(value.text)
        if shared_strings_map:
            return shared_strings_map.get(index, "")
        else:
            # shared-stringのテーブルから読み込む   
            return book.shared_strings.get_text(index)

def get_cell_datetime_value(element, time=None):
    if time is None: time = TARGET_TIME
    fl = float(element.v.text)
    return time.convert_to_datetime(fl)

def get_cell_time_value(element, time=None):
    if time is None: time = TARGET_TIME
    fl = float(element.v.text)
    return time.convert_to_time(fl)
    
def get_cell_number_format(element, book):
    style_index = element.s
    return book.style_sheet.get_format(style_index).number_format



"""
"""
class Cell(ElementProxy):
    def __init__(self, element, row, workbook):
        super(Cell, self).__init__(element)
        self._row = row
        self._book = workbook
        self._numfmt = None  # 書式バッファ
    
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
    
    def is_datetime(self):
        return self.is_number() and self.number_format.type == NUMVAL_TYPE_DATETIME
    
    def is_time(self):
        return self.is_number() and self.number_format.type == NUMVAL_TYPE_TIME
        
    def get_value(self):
        return get_cell_value(self._element, self._book)
    
    def clear_value(self):
        # とりあえず空文字列を代入。これでいいのか？
        self._element.t = ST_CellType.STR
        self._element.get_or_add_v().text = ""
        self._v = None
    
    def get_text(self, shared_strings_map=None):
        return get_cell_text(self._element, self._book, shared_strings_map)
    
    def _finish_shared_string(self, shared_strings, text):
        # shared_stringのインデックスを確定させる
        if self._element.t != ST_CellType.SHARED_STRING:
            return # さらに変更がおこり上書きされたので更新不要
        strid = shared_strings.add_string(text)
        self._element.get_or_add_v().text = str(strid)
    
    def get_datetime_value(self, time=None):
        return get_cell_datetime_value(self._element, time)

    def get_time_value(self, time=None):
        return get_cell_time_value(self._element, time)
    
    @property
    def raw(self):
        v = self._element.v
        if v is None:
            return None
        return v.text

    @property
    def value(self):
        return self.get_value()
    
    @property
    def text(self):
        return self.get_text()
    
    @text.setter
    def text(self, t):
        self._element.t = ST_CellType.SHARED_STRING
        text = self._element.get_or_add_v().text
        if text and text[0] == "M":
            self._book.shared_strings._set_pending_text(int(text[1:]), self, t)
        else:
            modid = self._book.shared_strings._set_pending_text(-1, self, t)
            self._element.v.text = "M{}".format(modid)

    @property
    def empty(self):
        return self._element.v is None

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
        style_index = self._element.s
        return self._book.style_sheet.get_format(style_index)
    
    @property
    def number_format(self):
        f = self._numfmt
        if f is None:
            f = self.style.number_format
            self._numfmt = f
        return f


class CellRange:
    """
    """
    def __init__(self, sheet, lefttop, rightbottom, iterbreak=True):
        self._sheet = sheet
        self._lt = lefttop
        self._rb = rightbottom
        self._iterbreak = iterbreak

    def __iter__(self):
        return self.cells
    
    @property
    def sheet(self):
        return self._sheet
    
    @property
    def coords(self):
        return (self._lt, self._rb)

    def row_cells(self):
        """
        行ごとにセルのリストを返す。
        Yields:
            List[Cell]:
        """
        brk = self._iterbreak
        r1, c1 = self._lt
        r2, c2 = self._rb
        for row in self._sheet.get_range_rows(r1, r2):
            cells = row.get_range_cells(c1, c2, emptynone=brk)
            if brk and cells is None:
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
    
    @property
    def cells(self):
        """
        すべてのセルを順に返す。
        """
        for cells in self.row_cells():
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
        List[Tuple[Str, Str]]: (テキスト、参照)のリスト　行の区別のない一次元のリスト
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
            ref = index_to_column(ci) + rowletter
            texts.append((text, ref))
    return texts
