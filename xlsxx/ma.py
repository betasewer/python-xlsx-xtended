from typing import List

from xlsxx.api import open_xlsx
from xlsxx.coord import ref_to_coord

from docxx.ma import OpcPackageFile

from xlsxx.parts.sml import SmlSheetMainPart
from xlsxx.proxy.workbook import Workbook
from xlsxx.proxy.sheet import Worksheet


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
    
    def read_v(self, start, tailcolumn, tailrow=-1, sequence=True):
        """ @method
        縦方向に値を読んで返す。
        Params:
            start(str): 開始セル参照
            tailcolumn(str): 終了列参照（含む）
            tailrow?(str): 終了行参照（含む）
        Returns:
            Tuple[Tuple[Str]]:
        """
        from xlsxx.proxy.sheet import read_sheet_vertical
        return read_sheet_vertical(self.cursheet(), start, tailcolumn, tailrow, sequence=sequence)

