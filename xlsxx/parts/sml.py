# encoding: utf-8

"""
|SmlSheetPart| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)
import os
import copy

from xlsxx.proxy.workbook import Workbook
from xlsxx.proxy.sheet import Worksheet
from xlsxx.proxy.shared_strings import SharedStrings
from xlsxx.proxy.styles import StyleSheet

from docxx.opc.constants import RELATIONSHIP_TYPE as RT, CONTENT_TYPE as CT
from docxx.oxml import parse_xml
from docxx.opc.part import XmlPart
from docxx.opc.packuri import PackURI
from docxx.shared import lazyproperty

#
def read_default_xml(filename):
    path = os.path.join(os.path.split(__file__)[0], '..', 'templates', filename)
    with open(path, 'rb') as f:
        xml_bytes = f.read()
    return xml_bytes


class SmlSheetMainPart(XmlPart):
    """
    """
    @property
    def core_properties(self):
        """
        A |CoreProperties| object providing read/write access to the core
        properties of this document.
        """
        return self.package.core_properties

    @property
    def workbook(self):
        """
        A |Document| object providing access to the content of this document.
        """
        return Workbook(self._element, self)
    
    @property
    def shared_strings(self):
        return self._shared_string_part.shared_strings
        
    @property
    def _shared_string_part(self):
        """
        """
        try:
            return self.part_related_by(RT.SHARED_STRINGS)
        except KeyError:
            ss_part = SmlSharedStringsPart.default(self.package)
            self.relate_to(ss_part, RT.SHARED_STRINGS)
            return ss_part
            
    @property
    def style_sheet(self):
        return self._styles_part.style_sheet
        
    @property
    def _styles_part(self):
        """
        """
        try:
            return self.part_related_by(RT.STYLES)
        except KeyError:
            ss_part = SmlStylesPart.default(self.package)
            self.relate_to(ss_part, RT.STYLES)
            return ss_part        
            
    def add_sheet_part(self):
        part = SmlWorkSheetPart.new(self.package)
        rId = self.relate_to(part, RT.WORKSHEET)
        return part, rId

    def save(self, path_or_stream):
        """
        Save this document to *path_or_stream*, which can be either a path to
        a filesystem location (a string) or a file-like object.
        """
        self.package.save(path_or_stream)
    
    def clone_document(self):
        """
        文書を複製する。
        """
        # opcパッケージをコピー
        from xlsxx.api import open_xlsx
        from docxx.opc.part import copy_part
        destpackage = open_xlsx().package
        copy_part(self.package, destpackage, destpackage)
        # 新しいワークブック
        return destpackage.main_document_part
    
    def fetch_textmap(self):
        """
        ワークブック内の文字列を全て取得する。
        """
        return self._shared_string_part.fetch_text()


#
class SmlWorkSheetPart(XmlPart): 
    @classmethod
    def new(cls, package):
        """Return newly created footer part."""
        partname = package.next_partname("/xl/worksheets/sheet%d.xml")
        content_type = CT.SML_WORKSHEET
        element = parse_xml(read_default_xml('default-worksheet.xml'))
        return cls(partname, content_type, element, package)

    def worksheet(self, workbook, sheetid):
        return Worksheet(self._element, self, workbook, sheetid)

#
class SmlSharedStringsPart(XmlPart):
    """
    Proxy for the sharedStrings.xml part.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._pending_cells = []

    @classmethod
    def new(cls, package, element):
        """
        """
        partname = PackURI('/xl/sharedStrings.xml')
        content_type = CT.SML_SHARED_STRINGS
        return cls(partname, content_type, element, package)

    @classmethod
    def default(cls, package):
        """
        Return a newly created styles part, containing a default set of
        elements.
        """
        element = parse_xml(read_default_xml("default-shared-strings.xml"))
        return cls.new(package, element)
    
    @property
    def shared_strings(self):
        return SharedStrings(self.element, self)
    
    def _set_pending_text(self, index, elcell, text):
        if index == -1:
            self._pending_cells.append((elcell, text))
            index = len(self._pending_cells)-1
        return index

    def _get_pending_text(self, index):
        elcell, text = self._pending_cells[index]
        return elcell, text

    def before_marshal(self):
        self.shared_strings._finish(self._pending_cells)
        self._pending_cells.clear()

    def fetch_text(self):
        texts = self.shared_strings.fetch_text()
        return texts


class SmlStylesPart(XmlPart):    
    @classmethod
    def new(cls, package, element):
        """
        """
        partname = PackURI('/word/styles.xml')
        content_type = CT.SML_STYLES
        return cls(partname, content_type, element, package)

    @classmethod
    def default(cls, package):
        """
        Return a newly created styles part, containing a default set of
        elements.
        """
        element = parse_xml(read_default_xml("styles.xml"))
        return cls.new(package, element)
        
    @property
    def style_sheet(self):
        return StyleSheet(self.element, self)
