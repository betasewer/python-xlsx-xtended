from typing import List

from xlsxx.api import open_xlsx
from xlsxx.coord import column_to_index, index_to_column, ref_to_coord

from docxx.ma import OpcPackageFile

from xlsxx.parts.sml import SmlSheetMainPart
from xlsxx.proxy.workbook import Workbook
from xlsxx.proxy.sheet import ReadingCells, Worksheet


class ExcelFile(OpcPackageFile):
    """ @type
    エクセルファイル。
    """
    def __init__(self, path=None, *, file=None):
        super().__init__(path=path, file=file)
        self._cursheetid = None

    @classmethod
    def new(cls, path):
        return cls(path, file=open_xlsx())

    def loadfile(self):
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
    
    def read_v(self, start, tailcolumn, tailrow=-1, sequential=True):
        """ @task
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
        return read_sheet_rows(self.cursheet(), start, tailcolumn, tailrow, readingdef=ReadingCells(sequential=sequential))

    def read_columns(self, start, columns, tailrow=-1, sequential=True):
        """ @task
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
        from xlsxx.proxy.sheet import read_sheet_rows
        startindex = ref_to_coord(start)[1]
        rdef = ReadingCells(sequential=sequential)
        for i, col in enumerate(columns, start=startindex):
            rdef.set_column_type(index_to_column(i), col)
        tailcolumn = startindex + len(columns) - 1
        return read_sheet_rows(self.cursheet(), start, tailcolumn, tailrow, readingdef=rdef)

    def write_v(self, start, rows, *, as_values=False):
        """ @task
        縦方向に値を書き込む。
        Params:
            start(str): 開始セル参照
            rows(Any): 行のリスト
        Returns:
            Tuple[Tuple[Str]]:
        """
        from xlsxx.coord import ref_to_coord
        srow, scol = ref_to_coord(start)

        sheet = self.cursheet()
        
        writings = sheet.writing_cells(as_values=as_values)
        for irow, row in enumerate(rows):
            for icol, val in enumerate(row):
                coord = (srow + irow, scol + icol)
                writings.add(coord, val)
        
        sheet.write_cells(writings)

    def write_cells(self, coord_text_dict, *, as_values=False):
        """ @task
        それぞれのセルに値を書き込む。
        Params:
            coord_text_dict(Any): 座標参照と値の組み合わせの辞書
        Returns:
            Tuple[Tuple[Str]]:
        """
        from xlsxx.coord import ref_to_coord

        sheet = self.cursheet()
        
        writings = sheet.writing_cells(as_values=as_values)
        for ref, val in coord_text_dict.items():
            coord = ref_to_coord(ref)
            writings.add(coord, val)
        
        sheet.write_cells(writings)
