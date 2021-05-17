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
class Font(ElementProxy):
    def __init__(self, element, ss):
        super().__init__(element)
        self._ss = ss
    
"""
"""
class Fill(ElementProxy):
    def __init__(self, element, ss):
        super().__init__(element)
        self._ss = ss
        
"""
"""
class CellStyle(ElementProxy):
    pass
    
"""
"""
class CellFormat(ElementProxy):
    pass

"""
"""
class StyleSheet(ElementProxy):
    __slots__ = ('_part', )

    def __init__(self, element, part):
        super().__init__(element)
        self._part = part

    @property
    def fonts(self):
        return [Font(x, self) for x in self._element.fonts_lst]
    
    @property
    def fills(self):
        return [Fill(x, self) for x in self._element.fills_lst]
    
    @property
    def styles(self):
        css = self._element.cellStyles_lst
        if len(css)==0: 
            return None
        return [CellStyle(x) for x in css[0].cellStyle_lst]
    
    @property
    def basic_formats(self):
        cfs = self._element.cellStyleXfs_lst
        if len(cfs)==0:
            return None
        return [CellFormat(i,x) for (i,x) in enumerate(cfs[0].xf)]
    
    @property
    def formats(self):
        cxs = self._element.cellXfs_lst
        if len(cxs)==0:
            return None
        return [CellFormat(i,x) for (i,x) in enumerate(cxs[0].xf)]
        
    @property
    def part(self):
        """
        The |DocumentPart| object of this document.
        """
        return self._part

