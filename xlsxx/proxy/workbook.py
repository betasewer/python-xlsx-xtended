# encoding: utf-8

"""
|Document| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docxx.shared import ElementProxy


"""
"""
class Workbook(ElementProxy):
    """
    Sequence [1..1]
        ssml:fileVersion [0..1]    File Version
        ssml:fileSharing [0..1]    File Sharing
        ssml:workbookPr [0..1]    Workbook Properties
        ssml:workbookProtection [0..1]    Workbook Protection
        ssml:bookViews [0..1]    Workbook Views
        ssml:sheets [1..1]    Sheets
        ssml:functionGroups [0..1]    Function Groups
        ssml:externalReferences [0..1]    External References
        ssml:definedNames [0..1]    Defined Names
        ssml:calcPr [0..1]    Calculation Properties
        ssml:oleSize [0..1]    OLE Size
        ssml:customWorkbookViews [0..1]    Custom Workbook Views
        ssml:pivotCaches [0..1]    PivotCaches
        ssml:smartTagPr [0..1]    Smart Tag Properties
        ssml:smartTagTypes [0..1]    Smart Tag Types
        ssml:webPublishing [0..1]    Web Publishing Properties
        ssml:fileRecoveryPr [0..*]    File Recovery Properties
        ssml:webPublishObjects [0..1]    Web Publish Objects
        ssml:extLst [0..1]    Future Feature Storage Area
    """

    __slots__ = ('_part', '_epoch')

    def __init__(self, element, part):
        super(Workbook, self).__init__(element)
        self._part = part
        self._epoch = None

    @property
    def core_properties(self):
        """
        A |CoreProperties| object providing read/write access to the core
        properties of this document.
        """
        return self._part.core_properties

    def _get_sheetentry(self, id=None, *, index=None, fallback=False):
        if id is not None:
            for elsheet in self._sheet_entries:
                if elsheet.sheetId == id:
                    return elsheet
        elif index is not None:
            if 0 <= index and index < len(self._sheet_entries):
                return self._sheet_entries[index]
        else:
            raise ValueError("Specify id or index")
        if not fallback:
            raise ValueError("No sheet entry found: id={}, index={}".format(id, index))
        return None
        
    def get_sheet_index(self, id=None):
        for i, elsheet in enumerate(self._sheet_entries):
            if id is not None and id == elsheet.sheetId:
                return i
        return None

    def sheet(self, id=None, *, index=None, entry=None):
        if entry is None:
            entry = self._get_sheetentry(id=id, index=index)
        wspart = self._part.related_parts.get(entry.rid)
        worksheet = wspart.worksheet(self, entry)
        return worksheet
    
    def add_sheet(self, *, name=None):
        part, rId = self._part.add_sheet_part()
        entry = self._element.sheets.add_sheet(rId, name)
        return part.worksheet(self, entry)
    
    def pick_sheet(self, index):
        """ インデックスが足りていなければシートを追加して返す """
        if index < len(self._sheet_entries):
            return self.sheet(index=index)
        else:
            return self.add_sheet()

    def delete_sheet(self, id=None, *, index=None, entry=None):
        """ 指定したシートを削除する """
        if len(self._sheet_entries) == 1:
            raise ValueError("シートは最低1つ必要です")
        if entry is None:
            entry = self._get_sheetentry(id=id, index=index)
        self._part.drop_rel(entry.rid) # シートのパーツを切り離す
        self._element.sheets.remove(entry) # シートのエントリーを削除する
    
    @property
    def topsheet(self):
        return self.sheet(index=0)
    
    @property
    def lastsheet(self):
        return self.sheet(entry=self._sheet_entries[-1])
    
    @property
    def sheets(self):
        return [self.sheet(entry=el) for el in self._sheet_entries]
    
    @property
    def _sheet_entries(self):
        return self._element.sheets.sheet_lst
    
    @property
    def shared_strings(self):
        return self._part.shared_strings
    
    @property
    def style_sheet(self):
        return self._part.style_sheet
    
    @property
    def part(self):
        """
        The |DocumentPart| object of this document.
        """
        return self._part

    def save(self, path_or_stream):
        """
        文書を保存する。
        Params:
            path_or_stream: ファイルパス、またはファイルオブジェクト
        """
        self._part.save(path_or_stream)
    
    def _copy_to(self, book, *, sheets=None):
        """ このワークブックの内容を別のワークブックにコピーする。 """
        pass

    @property
    def time_epoch(self):
        """
        時刻の初期値を取得する。
        """
        return self._part.get_time_epoch()

    @time_epoch.setter
    def time_epoch(self, name):
        """ """
        self._part.set_time_epoch(name)


