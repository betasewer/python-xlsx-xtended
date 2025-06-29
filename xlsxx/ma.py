from __future__ import annotations

from typing import List, Union, Any
from collections import defaultdict

from xlsxx.coord import column_to_index, index_to_column, ref_to_coord

from docxx.ma import OpcPackageFile

from xlsxx.parts.sml import SmlSheetMainPart
from xlsxx.proxy.workbook import Workbook
from xlsxx.proxy.sheet import ReadingCells, Worksheet, read_sheet_rows


class ExcelFile(OpcPackageFile):
    """ @type
    エクセルファイル。
    """
    def __init__(self, path=None, *, file=None):
        super().__init__(path=path, file=file)
        self._cursheetid = None

    @classmethod
    def new(cls, path):
        from xlsxx.api import open_xlsx
        return cls(path, file=open_xlsx())

    def loadfile(self):
        from xlsxx.api import open_xlsx
        return open_xlsx(self.pathstr)
    
    def savefile(self, p):
        self.load().save(p)
    
    def _with_path(self, p):
        return ExcelFile(p, file=self._file) # シートはリセットされる
    
    def topsheet(self) -> Worksheet:
        """ @method
        ワークブックの先頭のシートを返す。
        Returns:
            Any:
        """
        return self.workbook().topsheet
    
    def cursheet(self) -> Worksheet:
        """ @method
        現在選択されているシート。
        Returns:
            Any:
        """
        if self._cursheetid is None:
            return self.topsheet()
        return self.workbook().sheet(id=self._cursheetid)

    def sheets(self) -> List[Worksheet]:
        """ @method
        全てのシート。
        Returns:
            Any:
        """
        return self.workbook().sheets

    def with_sheet(self, index=None, *, id=None):
        """ @method
        シートを選択する。
        Params:
            index(int):
        """
        if id is not None:
            self._cursheetid = id
        elif index is not None:
            sh = self.workbook().sheet(index=index)
            self._cursheetid = sh.id

    def delete_sheet(self, index=None, *, id=None):
        """ シートを削除する。カレントシートであればカレントが無くなる。 """
        entry = self.workbook()._get_sheetentry(index=index, id=id)
        if id is None:
            id = entry.sheetId
        if self._cursheetid == id:
            self.workbook().delete_sheet(entry=entry)
            self._cursheetid = None
        else:
            self.workbook().delete_sheet(entry=entry)

    def workbook(self) -> Workbook:
        """ @method
        ワークブックを返す。
        Returns:
            Any:
        """
        return self.document().workbook
    
    def document(self) -> SmlSheetMainPart:
        """ @method
        ドキュメントを返す。
        Returns:
            Any: 
        """
        return self.load()
    
    @staticmethod
    def writing_v_table(start, **kwargs):
        return ExcelVTableWriting(start, **kwargs)

    @staticmethod
    def reading_cells(*, sequential=False, nofetch=False, as_values=False):
        return ReadingCells(sequential=sequential, nofetch=nofetch, as_values=as_values)
    
    def read_v(self, start, tailcolumn, tailrow=-1, sequential=True, as_values=False):
        """ @task nospirit
        縦方向に値を読んで返す。
        Params:
            start(str): 開始セル参照
            tailcolumn(str): 終了列参照（含む）
            tailrow?(str): 終了行参照（含む）
            sequential?(bool): 空の行があったら読み込みを止める
        Returns:
            Tuple[Tuple[Str]]:
        """
        from xlsxx.proxy.sheet import read_sheet_rows
        return read_sheet_rows(self.cursheet(), start, tailcolumn, tailrow, readingcells=ReadingCells(sequential=sequential, as_values=as_values))

    def read_columns(self, start, columns, tailrow=-1, sequential=True):
        """ @task nospirit
        値のタイプを指定して値を読んで返す。
        Params:
            start(str): 開始セル参照
            columns(Tuple[str]): 値タイプの並び
            tailrow?(str): 終了行参照（含む）
        Returns:
            Tuple[Tuple[Str]]:
        """
        if not columns:
            raise ValueError("カラム指定が空です")
        
        rdef = ReadingCells(sequential=sequential)
        startindex = ref_to_coord(start)[1]
        for i, col in enumerate(columns, start=startindex):
            rdef.set_column_type(index_to_column(i), col)
        
        tailcolumn = startindex + len(columns) - 1
        return read_sheet_rows(self.cursheet(), start, tailcolumn, tailrow, readingcells=rdef)

    def write(self, writings: 'ExcelWriting'):
        """ """
        writings.build_cell_formats(self.workbook())
        sheet = self.cursheet()
        sheet.write_cells(writings)

    def write_v(self, start, rows, *, rows_height=None, **writings_kwargs):
        """ @task nospirit
        縦方向に値を書き込む。
        Params:
            start(str): 開始セル参照
            rows(Any): 行のリスト
        Returns:
            Tuple[Tuple[Str]]:
        """
        writings = ExcelVTableWriting(start, **writings_kwargs)
        writings.rows(rows, rows_height=rows_height)
        self.write(writings)

    def time_epoch(self, name):
        """ @method
        グローバルな始原日付・時刻のシステムを設定する。
        Params:
            name(str): [excel|mac2008excel|libreoffice]
        Returns:
            Datetime:
        """
        self.workbook().time_epoch = name
    

#
#
#
class ExcelCellFormatWriting:
    def __init__(self):
        self.wraptext = False
        self.horizontal = "general"  # "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed"
        self.vertical = "top"        # "general" | "top" | "bottom" | "center" | "justify" | "distributed"
        self._xf_index = None

    def create_format(self, workbook: Workbook):
        """
        セルスタイルを作成する。
        Params:
            workbook(Workbook): ワークブックオブジェクト
        Returns:
            int: スタイルインデックス
        """
        xf = workbook.style_sheet.add_format()
        xf.alignment.wrap_text = self.wraptext
        xf.alignment.horizontal = self.horizontal 
        xf.alignment.vertical = self.vertical
        self._xf_index = xf.index
        return self._xf_index
    
    def get_xf_index(self):
        """
        セルスタイルのインデックスを取得する。
        Returns:
            int: セルスタイルのインデックス
        """
        return self._xf_index

#
#
class RowHeightFunctor:
    """
    列幅や行の高さを返す関数を作成する。
    列幅や行の高さは、列番号や行番号を受け取って値を返す。
    """
    def __init__(self, value, rows, start_offset=None, end_offset=None):
        if isinstance(value, (list, tuple)):
            self._fn = self.list_dim_functor(value)
        elif callable(value):
            self._fn = self.row_text_dim_functor(value, rows)
        else:
            raise TypeError("value must be list, tuple or callable")
        
        self.range = None
        if start_offset is not None and end_offset is not None:
            self.range = (start_offset, end_offset)

    def in_range(self, irowoffset):
        """
        行オフセットが範囲内かどうかを確認する。
        Params:
            irowoffset(int): 行オフセット
        Returns:
            bool: 範囲内ならTrue、そうでなければFalse
        """
        if self.range is None:
            return True
        return self.range[0] <= irowoffset and irowoffset < self.range[1]

    @staticmethod
    def list_dim_functor(widths):
        def functor(icol):
            if icol < 0 or icol >= len(widths):
                return None
            return widths[icol]
        return functor
    
    @staticmethod
    def row_text_dim_functor(f, rows):
        def functor(icol):
            if icol not in rows:
                return None
            return f([x[1] for x in rows[icol]])
        return functor
    
    def __call__(self, index):
        return self._fn(index)


#
#
#
class ExcelWriting:
    def __init__(self):
        self._rows = defaultdict(list)
        self._mincoord = None
        self._maxcoord = None

    def empty(self):
        return len(self._rows) == 0

    def add(self, coord, value):
        if not isinstance(coord, tuple):
            raise TypeError("coord must be tuple")
        row, col = coord
        self._rows[row].append((col, value))
        self._mincoord = (row, col) if self._mincoord is None else min(self._mincoord, (row, col)) 
        self._maxcoord = (row, col) if self._maxcoord is None else max(self._maxcoord, (row, col))

    def add_values(self, ref_value_dict: dict[str, Any]):
        """
        座標参照と値の辞書を受け取り、各座標に値を追加する。
        Params:
            ref_value_dict(dict): 座標参照文字と値の辞書
        """
        from xlsxx.coord import ref_to_coord
        for ref, val in ref_value_dict.items():
            coord = ref_to_coord(ref)
            self.add(coord, val)

    def mincoord(self):
        return self._mincoord
    
    def maxcoord(self):
        return self._maxcoord
    
    def get_rows(self):
        return self._rows.items()

    def has_row(self, index):
        return index in self._rows
    
    #
    #
    #
    def build_cell_formats(self, workbook: Workbook):
        ...
    
    def get_row_height(self, irowoffset: int) -> int:
        ...

    def get_column_width(self, icoloffset: int, _col) -> None:
        ...

    def columns_count(self) -> Union[int, None]:
        ...

    def cell_styles_bits(self, icoloffset: int):
        ...
    


class ExcelWritingColumn:
    def __init__(self, index):
        self.index = index
        self.width = None
        self.as_values = False
        self.cell_format = ExcelCellFormatWriting()

    def build_cell_format(self, workbook: Workbook):
        """
        スタイルをビルドする。
        Params:
            workbook(Workbook): ワークブックオブジェクト
        Returns:
            int: スタイルインデックス
        """
        return self.cell_format.create_format(workbook)


class ExcelVTableWriting(ExcelWriting):
    def __init__(self, start, **kwargs):
        super().__init__(**kwargs)
        from xlsxx.coord import ref_to_coord
        self._start = ref_to_coord(start)
        self._rows_heights = []
        self._cells_styles_bits = None
        self.cols: list[ExcelWritingColumn] = []
        self._rows_fill_offset = 0
    
    #
    #
    #
    def column(self,
        width=None,
        as_values=False,
        wraptext=False
    ):
        """
        列の設定を行う
        """
        icolstart = self._start[1]

        col = ExcelWritingColumn(index=icolstart + len(self.cols))
        col.width = width
        col.as_values = as_values
        col.cell_format.wraptext = wraptext

        self.cols.append(col)
        return self

    def rows(self, rows, rows_height=None):
        """
        行を流し込む。
        """
        srow, scol = self._start
        srow = srow + self._rows_fill_offset
        
        for irow, row in enumerate(rows):
            for icol, val in enumerate(row):
                coord = (srow + irow, scol + icol)
                self.add(coord, val)
        next_offset = self._rows_fill_offset + len(rows)

        if rows_height is not None:
            rh = RowHeightFunctor(rows_height, self._rows, start_offset=self._rows_fill_offset, end_offset=next_offset)
            self._rows_heights.append(rh)

        self._rows_fill_offset = next_offset
        return self

    def build_cell_formats(self, workbook: Workbook):
        """
        スタイルをビルドする。
        Params:
            workbook(Workbook): ワークブックオブジェクト
        Returns:
            None
        """
        for col in self.cols:
            col.build_cell_format(workbook)
        
        # 列内のセルは共通のスタイルを持つ
        bits = {}
        for col in self.cols:
            bits[col.index] = (
                col.cell_format.get_xf_index(),
                col.as_values
            )
        self._cells_styles_bits = bits

    #
    #
    #
    def get_row_height(self, irowoffset):
        """
        行の高さを返す。
        Params:
            irow(int): 行番号
        Returns:
            Callable: 行番号を受け取って行の高さを返す関数
        """
        if len(self._rows_heights) == 0:
            return None
        
        for row in self._rows_heights:
            if row.in_range(irowoffset):
                return row(irowoffset)
        return None
    
    def get_column(self, icoloffset):
        """
        列を返す。
        Params:
            icol(int): 列番号
        Returns:
            ExcelWritingColumn: 列オブジェクト
        """
        if len(self.cols) == 0:
            return None
        if icoloffset < 0 or icoloffset >= len(self.cols):
            return None
        return self.cols[icoloffset]
    
    def get_column_width(self, icoloffset, _col):
        """
        列の幅を返す。
        Params:
            icol(int): 列番号
        Returns:
            Callable: 列番号を受け取って列の幅を返す関数
        """
        if len(self.cols) == 0:
            return None
        if icoloffset < 0 or icoloffset >= len(self.cols):
            return None
        return self.cols[icoloffset].width

    def columns_count(self):
        """
        列の数が規定されていれば返す。
        Returns:
            int: 列の数
        """
        if len(self.cols) == 0:
            return None
        return len(self.cols)
    
    def cell_styles_bits(self, _irowoffset):
        """
        セルのスタイルをビットで返す。
        Returns:
            int: セルのスタイルを表すビット
        """
        return self._cells_styles_bits
