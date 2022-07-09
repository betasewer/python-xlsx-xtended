# encoding: utf-8

"""
|Document| and closely related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)
import re

from docxx.shared import ElementProxy


class NumberFormat(ElementProxy):
    """
    """
    def __init__(self, element):
        super().__init__(element)
    
    @property
    def id(self):
        return self._element.numFmtId

    @property
    def code(self):
        return self._element.formatCode    
        
    @property
    def type(self):
        return numfmt_lib.detect_value_type(self.code)


class Font(ElementProxy):
    """
    """
    def __init__(self, element):
        super().__init__(element)
    
    @property
    def name(self):
        return self._element._get_bool_val("name")

    @property
    def charset(self):
        return self._element._get_bool_val("charset")
    
    @property
    def family(self):
        return self._element._get_bool_val("family")
    
    @property
    def bold(self):
        return self._element._get_bool_val("b")
    
    @property
    def italic(self):
        return self._element._get_bool_val("i")
    
    @property
    def strike(self):
        return self._element._get_bool_val("strike")
    
    @property
    def outline(self):
        return self._element._get_bool_val("outline")
    
    @property
    def shadow(self):
        return self._element._get_bool_val("shadow")
    
    @property
    def condense(self):
        return self._element._get_bool_val("condense")
    
    @property
    def extend(self):
        return self._element._get_bool_val("extend")
    
    @property
    def underline(self):
        return self._element._get_bool_val("u")
    
    @property
    def size(self):
        return self._element.sz_val
    
    @property
    def color(self):
        return ColorFormat(self._element.color)
    

class ColorFormat(ElementProxy):
    def __init__(self, element):
        super().__init__(element)

    
"""
"""
class Fill(ElementProxy):
    def __init__(self, element):
        super().__init__(element)
        
        
"""
"""
class CellStyleName(ElementProxy):
    def __init__(self, element):
        super().__init__(element)
    
    
"""
"""
class CellFormat(ElementProxy):
    def __init__(self, element, ssheet):
        super().__init__(element)
        self._ss = ssheet

    @property
    def number_format(self):
        id = self._element.numFmtId
        if id is None:
            return None
        return self._ss.get_number_format_by_id(id)
    
    @property
    def number_value_type(self):
        id = self._element.numFmtId
        if id is None:
            return None
        fmt = self._ss.get_number_format_by_id(id)
        if fmt is not None:
            return numfmt_lib.detect_value_type(fmt.code)
        else:
            code = numfmt_lib.get_format(id)
            if code is None:
                return None
            return numfmt_lib.detect_value_type(code)
    
    @property
    def font(self):
        idx = self._element.fontId
        if idx is None:
            return None
        return self._ss.get_font(idx)

    @property
    def fill(self):
        self._element.fillId
    
    @property
    def border(self):
        self._element.borderId
    
    @property
    def basic_format(self):
        idx = self._element.xfId
        if idx is None:
            return None
        return self._ss.get_basic_format(idx)
    
    @property
    def quote_prefix(self):
        return self._element.quotePrefix
    
    @property
    def pivot_button(self):
        return self._element.pivotButton


"""
"""
class StyleSheet(ElementProxy):
    __slots__ = ('_part', )

    def __init__(self, element, part):
        super().__init__(element)
        self._part = part

    def get_number_format_by_id(self, id):
        numFmts = self._element.numFmts
        if numFmts is None:
            return None
        for elem in numFmts.numFmt_lst:
            nf = NumberFormat(elem)
            if nf.id == id:
                return nf
        return None
    
    def get_font(self, index):
        return Font(self._element.fonts_lst[index])
    
    def get_basic_format(self, index):
        return CellFormat(self._element.cellStyleXfs.xf_lst[index], self)
    
    def get_format(self, index):
        return CellFormat(self._element.cellXfs.xf_lst[index], self)
        
    @property
    def style_names(self):
        return [CellStyleName(x) for x in self._element.cellStyles.cellStyle_lst]
    
    @property
    def part(self):
        """
        The |DocumentPart| object of this document.
        """
        return self._part

#
#
# 数値フォーマット
# - 18.8.30 NumFmt
#
#
FORMAT_GENERAL = 0   # 0.8934 = 0.8934
FORMAT_NUMBER = 1    # 0.8934 = 0.8934
FORMAT_FLOAT = 2     # 0.8934 @ #,###.00 = .89
FORMAT_PERCENT = 3   # 0.8934 @ 0.00% = 89.34 % 
FORMAT_EXP = 4       # 0.8934 @ 0.00E+000 = 8.93E-001
FORMAT_FRACTION = 5  # 0.8934 @ # ??/?? = 67/75
FORMAT_CURRENCY = 6  # 0.8934 @ [$￥-411]#,##0;-[$￥-411]#,##0 = 1円
FORMAT_DATETIME = 10 # 18935.0560763889 @ YYYY/MM/DD H:MM = 1951/11/3 1:20:45
FORMAT_TIME = 11     # 0.0560763889 @ H:MM = 1:20:45
FORMAT_TEXT = 20     # そのまま

NUMVAL_TYPE_INT      = 0 # 値を整数として解釈
NUMVAL_TYPE_DATETIME = 1 # 値を日付＋時刻として解釈
NUMVAL_TYPE_TIME     = 2 # 値を時刻として解釈
NUMVAL_TYPE_FLOAT    = 3 # 値を小数として解釈

class NumFmtLibrary:
    def __init__(self):
        self._fmt_to_types = {}
        self._code_to_fmt  = {}
        self._chars = []
    
    def adds(self, *entries):
        for code, format, type in entries:
            self._fmt_to_types[format] = type
            self._code_to_fmt[code] = format
        return self
    
    def adds_type(self, *entries):
        for format, type in entries:
            self._fmt_to_types[format] = type

    def get_format(self, code):
        return self._code_to_fmt.get(code)
        
    def detect_std_format(self, format):
        if format in self._fmt_to_types:
            return self._fmt_to_types[format]

        format = self._make_valtype_detection_string(format)
        for stdfmt, type in self._fmt_to_types.items():
            if format == self._make_valtype_detection_string(stdfmt):
                return type

        return FORMAT_GENERAL
    
    def detect_value_type(self, format):
        if format in self._fmt_to_types:
            type = self._fmt_to_types[format]
            if type == FORMAT_DATETIME:
                return NUMVAL_TYPE_DATETIME
            elif type == FORMAT_TIME:
                return NUMVAL_TYPE_TIME
            elif type in (FORMAT_FLOAT, FORMAT_PERCENT, FORMAT_EXP, FORMAT_FRACTION, FORMAT_CURRENCY):
                return NUMVAL_TYPE_FLOAT
            else:
                return NUMVAL_TYPE_INT
            
        format = self.make_valtype_detection_string(format)
        if any(x in format for x in ('hh', 'h', 'mm:', 'm:', 'ss', 's')):
            return NUMVAL_TYPE_TIME
        if any(x in format for x in ('yyyy', 'yy', 'ggge', 'ge', 'mm', 'm', 'dd', 'd', 'ww')):
            return NUMVAL_TYPE_DATETIME

        if any(x in format for x in (".","%","e","/")):
            # 小数点・パーセント・指数表記・分数
            return NUMVAL_TYPE_FLOAT
        
        return NUMVAL_TYPE_INT
    
    def make_valtype_detection_string(self, fmt):
        fmt = fmt.lower()
        fmt = re.sub(r'''(\"[^\"]*?\")''', lambda m:" ", fmt) # ""でくくられたリテラルを消去する
        fmt = re.sub(r'''(\[[^\]]*?\])''', lambda m:" ", fmt) # []でくくられた色指定？を消去する
        return fmt


numfmt_lib = NumFmtLibrary()
numfmt_lib.adds(
    (0, 'General', FORMAT_GENERAL),
    (0, 'Standard', FORMAT_GENERAL),
    (1, '0', FORMAT_NUMBER),
    (2, '0.00', FORMAT_FLOAT),
    (3, '#,##0', FORMAT_NUMBER),
    (4, '#,##0.00', FORMAT_FLOAT),
    (9, '0%', FORMAT_PERCENT),
    (10, '0.00%', FORMAT_PERCENT),
    (11, '0.00E+00', FORMAT_EXP),
    (12, '# ?/?', FORMAT_FRACTION),
    (13, '# ??/??', FORMAT_FRACTION),
    (14, 'mm-dd-yy', FORMAT_DATETIME),
    (15, 'd-mmm-yy', FORMAT_DATETIME),
    (16, 'd-mmm', FORMAT_DATETIME),
    (17, 'mmm-yy', FORMAT_DATETIME),
    (18, 'h:mm AM/PM', FORMAT_TIME),
    (19, 'h:mm:ss AM/PM', FORMAT_TIME),
    (20, 'h:mm', FORMAT_TIME),
    (21, 'h:mm:ss', FORMAT_TIME),
    (22, 'm/d/yy h:mm', FORMAT_DATETIME),
    (30, 'm/d/yy', FORMAT_DATETIME),
    (37, '#,##0 ;(#,##0)', FORMAT_NUMBER),
    (38, '#,##0 ;[Red](#,##0)', FORMAT_NUMBER),
    (39, '#,##0.00;(#,##0.00)', FORMAT_NUMBER),
    (40, '#,##0.00;[Red](#,##0.00)', FORMAT_NUMBER),
    (45, 'mm:ss', FORMAT_TIME),
    (46, '[h]:mm:ss', FORMAT_TIME),
    (47, 'mmss.0', FORMAT_TIME),
    (48, '##0.0E+0', FORMAT_EXP),
    (49, '@', FORMAT_TEXT),
)
# Japanese Specific
numfmt_lib.adds(
    (27, '[$-411]ge.m.d', FORMAT_DATETIME),
    (36, '[$-411]ge.m.d', FORMAT_DATETIME),
    (50, '[$-411]ge.m.d', FORMAT_DATETIME),
    (57, '[$-411]ge.m.d', FORMAT_DATETIME),
    (28, '[$-411]ggge"年"m"月"d"日"', FORMAT_DATETIME),
    (29, '[$-411]ggge"年"m"月"d"日"', FORMAT_DATETIME),
    (51, '[$-411]ggge"年"m"月"d"日"', FORMAT_DATETIME),
    (54, '[$-411]ggge"年"m"月"d"日"', FORMAT_DATETIME),
    (58, '[$-411]ggge"年"m"月"d"日"', FORMAT_DATETIME),
    (31, 'yyyy"年"m"月"d"日"', FORMAT_DATETIME),
    (34, 'yyyy"年"m"月"', FORMAT_DATETIME),
    (52, 'yyyy"年"m"月"', FORMAT_DATETIME),
    (55, 'yyyy"年"m"月"', FORMAT_DATETIME),
    (32, 'h"時"mm"分"', FORMAT_DATETIME),
    (33, 'h"時"mm"分"ss"秒"', FORMAT_DATETIME),
    (35, 'm"月"d"日"', FORMAT_DATETIME),
    (53, 'm"月"d"日"', FORMAT_DATETIME),
    (56, 'm"月"d"日"', FORMAT_DATETIME),
)

