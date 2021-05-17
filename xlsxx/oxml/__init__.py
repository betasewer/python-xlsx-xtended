"""
#
#
#
#         XLSX / spread sheet ml
#
#
#
"""
#
#
#
from .workbook import (
    CT_Workbook, CT_Sheets, CT_Sheet
)
from .sheet import (
    CT_Worksheet, CT_SheetData, CT_SheetDimension, CT_Cols, CT_Col, CT_Row, CT_Cell, 
    CT_MergeCells, CT_MergeCell, CT_Cell, CT_CellFormula,
)
from .shared_strings import (
    CT_SST, CT_RST, CT_PhoneticPr, CT_PhoneticRun, 
    CT_RElt, CT_RPrElt
)
from .styles import (
    CT_BooleanProperty, CT_Border, CT_BorderPr, CT_Borders, 
    CT_CellAlignment, CT_CellProtection, CT_CellStyle, CT_CellStyleXfs, CT_CellStyles, CT_CellXfs, 
    CT_Color, CT_Colors, CT_Dxf, CT_Dxfs, CT_Fill, CT_Fills, 
    CT_Font, CT_FontName, CT_FontScheme, CT_FontSize, CT_Fonts, 
    CT_GradientFill, CT_GradientStop, CT_IndexedColors, CT_IntProperty, CT_MRUColors, 
    CT_NumFmt, CT_NumFmts, CT_PatternFill, CT_RgbColor, CT_Stylesheet, CT_TableStyle, CT_TableStyleElement, 
    CT_TableStyles, CT_UnderlineProperty, CT_VerticalAlignFontProperty, CT_Xf,
)
from .shared import (
    CT_Xstring, CT_XStringElement,
    CT_Extension, CT_ExtensionList,
)

from docxx.oxml import register_element_cls, extend_nsmap, parse_xml
extend_nsmap({
    "ssml" : "http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
})
parse_xml = parse_xml # エイリアス


#
register_element_cls('ssml:workbook',  CT_Workbook)
register_element_cls('ssml:sheets',    CT_Sheets)
register_element_cls('ssml:sheet',     CT_Sheet)
register_element_cls('ssml:ext',       CT_Extension)
register_element_cls('ssml:extLst',    CT_ExtensionList)
#
register_element_cls('ssml:worksheet',  CT_Worksheet)
register_element_cls('ssml:dimension',  CT_SheetDimension)
register_element_cls('ssml:sheetData',  CT_SheetData)
register_element_cls('ssml:cols',       CT_Cols)
register_element_cls('ssml:col',        CT_Col)
register_element_cls('ssml:row',        CT_Row)
register_element_cls('ssml:c',          CT_Cell)
register_element_cls('ssml:f',          CT_CellFormula)
register_element_cls('ssml:is',         CT_RST)
register_element_cls('ssml:v',          CT_Xstring)
register_element_cls('ssml:mergeCells', CT_MergeCells)
register_element_cls('ssml:mergeCell',  CT_MergeCell)
#
register_element_cls('ssml:sst',        CT_SST)
register_element_cls('ssml:si',         CT_RST)
register_element_cls('ssml:t',          CT_Xstring)
register_element_cls('ssml:rPr',        CT_RPrElt)
register_element_cls('ssml:r',          CT_RElt)
register_element_cls('ssml:rPh',        CT_PhoneticRun)
register_element_cls('ssml:phoneticPr', CT_PhoneticPr)
#
register_element_cls('ssml:styleSheet', CT_Stylesheet)
register_element_cls('ssml:numFmts', CT_NumFmts)
register_element_cls('ssml:fonts', CT_Fonts)
register_element_cls('ssml:fills', CT_Fills)
register_element_cls('ssml:borders', CT_Borders)
register_element_cls('ssml:cellStyleXfs', CT_CellStyleXfs)
register_element_cls('ssml:cellXfs', CT_CellXfs)
register_element_cls('ssml:cellStyles', CT_CellStyles)
register_element_cls('ssml:dxfs', CT_Dxfs)
register_element_cls('ssml:tableStyles', CT_TableStyles)
register_element_cls('ssml:colors', CT_Colors)
register_element_cls('ssml:extLst', CT_ExtensionList)
register_element_cls('ssml:border', CT_Border)
register_element_cls('ssml:left', CT_BorderPr)
register_element_cls('ssml:right', CT_BorderPr)
register_element_cls('ssml:top', CT_BorderPr)
register_element_cls('ssml:bottom', CT_BorderPr)
register_element_cls('ssml:diagonal', CT_BorderPr)
register_element_cls('ssml:vertical', CT_BorderPr)
register_element_cls('ssml:horizontal', CT_BorderPr)
register_element_cls('ssml:color', CT_Color)
register_element_cls('ssml:font', CT_Font)
register_element_cls('ssml:fill', CT_Fill)
register_element_cls('ssml:patternFill', CT_PatternFill)
register_element_cls('ssml:gradientFill', CT_GradientFill)
register_element_cls('ssml:fgColor', CT_Color)
register_element_cls('ssml:bgColor', CT_Color)
register_element_cls('ssml:stop', CT_GradientStop)
register_element_cls('ssml:color', CT_Color)
register_element_cls('ssml:numFmt', CT_NumFmt)
register_element_cls('ssml:xf', CT_Xf)
register_element_cls('ssml:xf', CT_Xf)
register_element_cls('ssml:alignment', CT_CellAlignment)
register_element_cls('ssml:protection', CT_CellProtection)
register_element_cls('ssml:extLst', CT_ExtensionList)
register_element_cls('ssml:cellStyle', CT_CellStyle)
register_element_cls('ssml:extLst', CT_ExtensionList)
register_element_cls('ssml:dxf', CT_Dxf)
register_element_cls('ssml:font', CT_Font)
register_element_cls('ssml:numFmt', CT_NumFmt)
register_element_cls('ssml:fill', CT_Fill)
register_element_cls('ssml:alignment', CT_CellAlignment)
register_element_cls('ssml:border', CT_Border)
register_element_cls('ssml:protection', CT_CellProtection)
register_element_cls('ssml:extLst', CT_ExtensionList)
register_element_cls('ssml:indexedColors', CT_IndexedColors)
register_element_cls('ssml:mruColors', CT_MRUColors)
register_element_cls('ssml:rgbColor', CT_RgbColor)
register_element_cls('ssml:color', CT_Color)
register_element_cls('ssml:tableStyle', CT_TableStyle)
register_element_cls('ssml:tableStyleElement', CT_TableStyleElement)
register_element_cls('ssml:name', CT_FontName)
register_element_cls('ssml:charset', CT_IntProperty)
register_element_cls('ssml:family', CT_IntProperty)
register_element_cls('ssml:b', CT_BooleanProperty)
register_element_cls('ssml:i', CT_BooleanProperty)
register_element_cls('ssml:strike', CT_BooleanProperty)
register_element_cls('ssml:outline', CT_BooleanProperty)
register_element_cls('ssml:shadow', CT_BooleanProperty)
register_element_cls('ssml:condense', CT_BooleanProperty)
register_element_cls('ssml:extend', CT_BooleanProperty)
register_element_cls('ssml:color', CT_Color)
register_element_cls('ssml:sz', CT_FontSize)
register_element_cls('ssml:u', CT_UnderlineProperty)
register_element_cls('ssml:vertAlign', CT_VerticalAlignFontProperty)
register_element_cls('ssml:scheme', CT_FontScheme)
