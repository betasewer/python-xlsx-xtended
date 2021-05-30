
from docxx.shared import ElementProxy, AttributeProperty
from xlsxx.oxml.simpletypes import ST_CellType
from xlsxx.coord import ref_to_coord, coord_to_ref, range_ref_to_coord, column_to_index, split_ref, get_range_coord

"""
"""
class CellRow(ElementProxy):
    def __init__(self, element, workbook):
        super(CellRow, self).__init__(element)
        self._workbook = workbook
    
    @property
    def ref(self):
        return self._element.ref
    
    def cells(self):
        for elcell in self._element.cell_lst:
            yield Cell(elcell, self._element, self._workbook)
        
    def get_range_cells(self, head, tail):
        """
        範囲内のセルを取得する。
        Params:
            head(int): 先頭インデックス
            tail(int): 最終インデックス（含む）
        Returns:
            List[Cell]:
        """
        return [Cell(x, self._element, self._workbook) for x in self._element.cell_lst[head:tail+1]]
        
    def cell(self, column=0):
        elcell = self._element.cell_lst[column]
        return Cell(elcell, self._element, self._workbook)

    def new_empties_until(self, column):
        """ 空のセルを複数追加する。
        Params:
            columnindex(int): カラム0座標
        """
        cur = len(self._element.cell_lst)-1
        if cur >= column:
            return
        for i in range(column-cur):
            self.add_cell(cur+i+1)

    def add_cell(self, column):
        """ 空のセルを追加する。
        Params:
            columnindex(int): カラム0座標
        """
        cell = self._element._add_cell()
        cell.ref = coord_to_ref((self.ref-1, column))

"""
"""
class Cell(ElementProxy):
    def __init__(self, element, row, workbook):
        super(Cell, self).__init__(element)
        self._row = row
        self._book = workbook
        self._t = None # 内部テキストバッファ
    
    def is_string(self):
        return self._element.type in (
            ST_CellType.shared_string, ST_CellType.string, ST_CellType.inline_string
        )
    
    def is_number(self):
        return self._element.type == ST_CellType.number
    
    def is_boolean(self):
        return self._element.type == ST_CellType.boolean
    
    def is_error(self):
        return self._element.type == ST_CellType.error
    
    @property
    def value(self):
        v = self._element.value
        if self.is_string():
            return v.text
        else:
            # TODO: それぞれの型の変換処理を行う
            return v.text
    
    def clear_value(self):
        # とりあえず空文字列を代入。これでいいのか？
        self._element.type = ST_CellType.string
        self._element.get_or_add_value().text = ""
    
    @property
    def text(self):
        if self._t:
            return self._t
        elif self._element.type == ST_CellType.shared_string:
            index = int(self.value)
            si = self._book.shared_strings.get_item(index)
            if si is None:
                return ""
            self._t = si.text
            return si.text
        else:
            v = self._element.value
            if v is None:
                return ""
            return v.text
    
    @text.setter
    def text(self, t):
        self._element.type = ST_CellType.shared_string
        self._book.shared_strings._add_pending_cell(self)
        self._t = t
        if self._element.value is not None:
            self._element.value.text = ""
    
    def _finish_shared_string(self, shared_strings):
        # shared_stringのインデックスを確定させる
        if self._t is None:
            raise ValueError("shared_stringのインデックスが不正です")
        strid = shared_strings.add_string(self._t)
        self._element.get_or_add_value().text = str(strid)

    @property
    def type(self):
        return self._element.type

    @property
    def empty(self):
        return self.value is None

    @property
    def ref(self):
        return self._element.ref

    @property
    def coord(self):
        return ref_to_coord(self._element.ref)

    @property
    def row(self):
        return self.coord[0]

    @property
    def column(self):
        return self.coord[1]
    
    @property
    def column_letter(self):
        return split_ref(self._element.ref)[0]
    

class CellRange:
    """
    """
    def __init__(self, sheet, lefttop, rightbottom, orientation=None):
        self._sheet = sheet
        self._lt = lefttop
        self._rb = rightbottom
        self._orientation = orientation

    def __iter__(self):
        return self.cells()
    
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
        r1, c1 = self._lt
        r2, c2 = self._rb
        for row in self._sheet.get_range_rows(r1, r2):
            cells = row.get_range_cells(c1, c2)
            if not cells:
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
    
    def cells(self):
        """
        すべてのセルを順に返す。
        """
        for cells in self.row_cells():
            for cell in cells:
                yield cell
    

    