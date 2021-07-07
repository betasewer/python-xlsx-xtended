
from xlsxx.api import open_xlsx
from xlsxx.coord import ref_to_coord

from machaon.types.file import BasicLoadFile


class ExcelFile(BasicLoadFile):
    """ @type
    エクセルファイル。
    """
    def __init__(self, path=None, *, file=None):
        super().__init__(path=path, file=file)
        self._cursheetid = None

    def loadfile(self):
        return open_xlsx(self.pathstr)
    
    def savefile(self, p):
        self.load().save(p)
    
    def _with_path(self, p):
        return ExcelFile(p, file=self._file) # シートはリセットされる
    
    def topsheet(self):
        """ @method
        ワークブックの先頭のシートを返す。
        Returns:
            Any:
        """
        return self.load().workbook.topsheet
    
    def cursheet(self):
        """ @method
        現在選択されているシート。
        Returns:
            Any:
        """
        if self._cursheetid is None:
            return self.topsheet()
        return self.load().workbook.sheet(id=self._cursheetid)

    def with_sheet(self, index):
        """ @method
        シートを選択する。
        Params:
            index(int):
        """
        sh = self.load().workbook.sheet(index=index)
        self._cursheetid = sh.id

    def workbook(self):
        """ @method
        ワークブックを返す。
        Returns:
            Any:
        """
        return self.load().workbook
    
    def document(self):
        """ @method
        ドキュメントを返す。
        Returns:
            Any: 
        """
        return self.load()
        
