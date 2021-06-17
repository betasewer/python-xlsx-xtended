# encoding: utf-8

"""
|Document| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docxx.shared import ElementProxy, AttributeProperty

"""
"""
class SharedStrings(ElementProxy):
    """
    """

    __slots__ = ('_part', )

    def __init__(self, element, part):
        super(SharedStrings, self).__init__(element)
        self._part = part
    
    @property
    def items(self):
        """        """
        return [StringItem(x) for x in self._element.si_lst]
    
    def get_item(self, index):
        sis = self.items
        if 0<=index and index<len(sis):
            return sis[index]
        return None

    def add_string(self, string):
        si = StringItem(self._element._add_si())
        si.text = string
        newid = len(self._element.si_lst)-1
        return newid
    
    def _add_pending_text(self, cell):
        return self._part._add_pending_text(cell)
    
    def _get_pending_text(self, id):
        return self._part._get_pending_text(id)

    def _finish_before_marshal(self, pending_cells):
        # 新たに書き込まれたテキストを格納する
        for cell in pending_cells:
            cell._finish_shared_string(self)
        
        c = len(self._element.si_lst)
        self._element.count = c
        self._element.uniqueCount = c


class StringItem(ElementProxy):
    """
    SharedStringsの要素。
    """
    @property
    def text(self):
        t = self._element.t
        if t is None:
            if self.has_runs():
                return "".join(x.text for x in self._runs)
            return ""
        return t.text
    
    @text.setter
    def text(self, text):
        elem = self._element.get_or_add_t()
        elem.text = text

    def has_runs(self):
        return self._element.r_lst is not None
    
    @property
    def _runs(self):
        return (Run(r) for r in self._element.r_lst)


class Run(ElementProxy):
    """    
    書式付きのテキスト。
    """
    def __init__(self, element):
        super(Run, self).__init__(element)
        self._rPr = element.get_or_add_rPr()
        """ rPr elements
        ssml:rFont [0..1]    Font
        ssml:charset [0..1]    Character Set
        ssml:family [0..1]    Font Family
        ssml:b [0..1]    Bold
        ssml:i [0..1]    Italic
        ssml:strike [0..1]    Strike Through
        ssml:outline [0..1]    Outline
        ssml:shadow [0..1]    Shadow
        ssml:condense [0..1]    Condense
        ssml:extend [0..1]    Extend
        ssml:color [0..1]    Text Color
        ssml:sz [0..1]    Font Size
        ssml:u [0..1]    Underline
        ssml:vertAlign [0..1]    Vertical Alignment
        ssml:scheme [0..1]    Font Scheme
        """
    
    @property
    def text(self):
        return self.element.t.text
    
