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
    CT_Workbook, CT_Sheets, CT_Sheet, 
    CT_BookView, CT_BookViews, CT_CalcPr, CT_CustomWorkbookView, CT_CustomWorkbookViews, 
    CT_DefinedName, CT_DefinedNames, CT_ExternalReference, CT_ExternalReferences, 
    CT_FileRecoveryPr, CT_FileSharing, CT_FileVersion, CT_FunctionGroup, CT_FunctionGroups, 
    CT_OleSize, CT_PivotCache, CT_PivotCaches, CT_SmartTagPr, CT_SmartTagType, 
    CT_SmartTagTypes, CT_WebPublishObject, CT_WebPublishObjects, CT_WebPublishing, 
    CT_WorkbookPr, CT_WorkbookProtection,
    CT_SheetBackgroundPicture
)
from .sheet import (
    CT_Worksheet, CT_Dialogsheet, CT_Chartsheet, CT_SheetData, CT_SheetDimension, CT_Cols, CT_Col, CT_Row, CT_Cell, 
    CT_MergeCells, CT_MergeCell, CT_Cell, CT_CellFormula,     
    CT_Break, CT_CellSmartTag, CT_CellSmartTagPr, CT_CellSmartTags, CT_CellWatch, 
    CT_CellWatches, CT_CfRule, CT_Cfvo, CT_ColorScale, CT_ConditionalFormatting, 
    CT_Control, CT_Controls, CT_CustomProperties, CT_CustomProperty, CT_CustomSheetView, 
    CT_CustomSheetViews, CT_DataBar, CT_DataConsolidate, CT_DataRef, CT_DataRefs, CT_DataValidation, 
    CT_DataValidations, CT_Drawing, CT_HeaderFooter, CT_Hyperlink, CT_Hyperlinks, 
    CT_IconSet, CT_IgnoredError, CT_IgnoredErrors, CT_InputCells, CT_LegacyDrawing, CT_OleObject, 
    CT_OleObjects, CT_OutlinePr, CT_PageBreak, CT_PageMargins, CT_PageSetUpPr, CT_PageSetup, CT_Pane, 
    CT_PivotSelection, CT_PrintOptions, CT_ProtectedRange, CT_ProtectedRanges, 
    CT_Scenario, CT_Scenarios, CT_Selection, CT_SheetCalcPr, 
    CT_SheetFormatPr, CT_SheetPr, CT_SheetProtection, CT_SheetView, CT_SheetViews, CT_SmartTags, 
    CT_TablePart, CT_TableParts, CT_WebPublishItem, CT_WebPublishItems,
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
    CT_TableStyles, CT_UnderlineProperty, CT_VerticalAlignFontProperty, CT_Xf
)
from .auto_filter import (
    CT_AutoFilter, CT_SortCondition, CT_ColorFilter, CT_CustomFilter, CT_CustomFilters, 
    CT_SortState, CT_DateGroupItem, CT_DynamicFilter, CT_Filter, CT_FilterColumn, CT_Filters,
    CT_IconFilter, CT_Top10
)
from .pivot_table import (
    CT_PivotArea, CT_PivotAreaReferences, CT_PivotAreaReference, CT_Index
)
from .shared import (
    CT_Xstring, CT_XStringElement,
    CT_Extension, CT_ExtensionList
)

from docxx.oxml import register_element_cls, extend_nsmap, parse_xml
extend_nsmap({
    "ssml" : "http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
})
parse_xml = parse_xml # エイリアス


# sheet
register_element_cls('ssml:worksheet', CT_Worksheet)
register_element_cls('ssml:chartsheet', CT_Chartsheet)
register_element_cls('ssml:dialogsheet', CT_Dialogsheet)
register_element_cls('ssml:sheetPr', CT_SheetPr)
register_element_cls('ssml:dimension', CT_SheetDimension)
register_element_cls('ssml:sheetViews', CT_SheetViews)
register_element_cls('ssml:sheetFormatPr', CT_SheetFormatPr)
register_element_cls('ssml:cols', CT_Cols)
register_element_cls('ssml:sheetData', CT_SheetData)
register_element_cls('ssml:sheetProtection', CT_SheetProtection)
register_element_cls('ssml:autoFilter', CT_AutoFilter)
register_element_cls('ssml:sortState', CT_SortState)
register_element_cls('ssml:dataConsolidate', CT_DataConsolidate)
register_element_cls('ssml:customSheetViews', CT_CustomSheetViews)
register_element_cls('ssml:phoneticPr', CT_PhoneticPr)
register_element_cls('ssml:conditionalFormatting', CT_ConditionalFormatting)
register_element_cls('ssml:printOptions', CT_PrintOptions)
register_element_cls('ssml:pageMargins', CT_PageMargins)
register_element_cls('ssml:pageSetup', CT_PageSetup)
register_element_cls('ssml:headerFooter', CT_HeaderFooter)
register_element_cls('ssml:rowBreaks', CT_PageBreak)
register_element_cls('ssml:colBreaks', CT_PageBreak)
register_element_cls('ssml:customProperties', CT_CustomProperties)
register_element_cls('ssml:drawing', CT_Drawing)
register_element_cls('ssml:legacyDrawing', CT_LegacyDrawing)
register_element_cls('ssml:legacyDrawingHF', CT_LegacyDrawing)
register_element_cls('ssml:picture', CT_SheetBackgroundPicture)
register_element_cls('ssml:oleObjects', CT_OleObjects)
register_element_cls('ssml:extLst', CT_ExtensionList)
register_element_cls('ssml:sheetCalcPr', CT_SheetCalcPr)
register_element_cls('ssml:protectedRanges', CT_ProtectedRanges)
register_element_cls('ssml:scenarios', CT_Scenarios)
register_element_cls('ssml:mergeCells', CT_MergeCells)
register_element_cls('ssml:dataValidations', CT_DataValidations)
register_element_cls('ssml:hyperlinks', CT_Hyperlinks)
register_element_cls('ssml:cellWatches', CT_CellWatches)
register_element_cls('ssml:ignoredErrors', CT_IgnoredErrors)
register_element_cls('ssml:smartTags', CT_SmartTags)
register_element_cls('ssml:controls', CT_Controls)
register_element_cls('ssml:webPublishItems', CT_WebPublishItems)
register_element_cls('ssml:tableParts', CT_TableParts)
register_element_cls('ssml:row', CT_Row)
register_element_cls('ssml:col', CT_Col)
register_element_cls('ssml:c', CT_Cell)
register_element_cls('ssml:f', CT_CellFormula)
register_element_cls('ssml:v', CT_Xstring)
register_element_cls('ssml:tabColor', CT_Color)
register_element_cls('ssml:outlinePr', CT_OutlinePr)
register_element_cls('ssml:pageSetUpPr', CT_PageSetUpPr)
register_element_cls('ssml:sheetView', CT_SheetView)
register_element_cls('ssml:pane', CT_Pane)
register_element_cls('ssml:selection', CT_Selection)
register_element_cls('ssml:pivotSelection', CT_PivotSelection)
register_element_cls('ssml:pivotArea', CT_PivotArea)
register_element_cls('ssml:brk', CT_Break)
register_element_cls('ssml:dataRefs', CT_DataRefs)
register_element_cls('ssml:dataRef', CT_DataRef)
register_element_cls('ssml:mergeCell', CT_MergeCell)
register_element_cls('ssml:cellSmartTags', CT_CellSmartTags)
register_element_cls('ssml:cellSmartTag', CT_CellSmartTag)
register_element_cls('ssml:cellSmartTagPr', CT_CellSmartTagPr)
register_element_cls('ssml:customSheetView', CT_CustomSheetView)
register_element_cls('ssml:dataValidation', CT_DataValidation)
#register_element_cls('ssml:formula1', ST_Formula) ElementBaseを継承していないので登録できない
#register_element_cls('ssml:formula2', ST_Formula) 
#register_element_cls('ssml:formula', ST_Formula)
register_element_cls('ssml:cfRule', CT_CfRule)
register_element_cls('ssml:colorScale', CT_ColorScale)
register_element_cls('ssml:dataBar', CT_DataBar)
register_element_cls('ssml:iconSet', CT_IconSet)
register_element_cls('ssml:hyperlink', CT_Hyperlink)
register_element_cls('ssml:cfvo', CT_Cfvo)
register_element_cls('ssml:color', CT_Color)
register_element_cls('ssml:oddHeader', CT_Xstring)
register_element_cls('ssml:oddFooter', CT_Xstring)
register_element_cls('ssml:evenHeader', CT_Xstring)
register_element_cls('ssml:evenFooter', CT_Xstring)
register_element_cls('ssml:firstHeader', CT_Xstring)
register_element_cls('ssml:firstFooter', CT_Xstring)
register_element_cls('ssml:scenario', CT_Scenario)
register_element_cls('ssml:protectedRange', CT_ProtectedRange)
register_element_cls('ssml:inputCells', CT_InputCells)
register_element_cls('ssml:cellWatch', CT_CellWatch)
register_element_cls('ssml:customPr', CT_CustomProperty)
register_element_cls('ssml:oleObject', CT_OleObject)
register_element_cls('ssml:webPublishItem', CT_WebPublishItem)
register_element_cls('ssml:control', CT_Control)
register_element_cls('ssml:ignoredError', CT_IgnoredError)
register_element_cls('ssml:tablePart', CT_TablePart)

# shared-string
register_element_cls('ssml:sst',        CT_SST)
register_element_cls('ssml:si',         CT_RST)
register_element_cls('ssml:t',          CT_Xstring)
register_element_cls('ssml:rPr',        CT_RPrElt)
register_element_cls('ssml:r',          CT_RElt)
register_element_cls('ssml:rPh',        CT_PhoneticRun)
register_element_cls('ssml:phoneticPr', CT_PhoneticPr)

# styles
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

# workbook
register_element_cls('ssml:workbook', CT_Workbook)
register_element_cls('ssml:fileVersion', CT_FileVersion)
register_element_cls('ssml:fileSharing', CT_FileSharing)
register_element_cls('ssml:workbookPr', CT_WorkbookPr)
register_element_cls('ssml:workbookProtection', CT_WorkbookProtection)
register_element_cls('ssml:bookViews', CT_BookViews)
register_element_cls('ssml:sheets', CT_Sheets)
register_element_cls('ssml:functionGroups', CT_FunctionGroups)
register_element_cls('ssml:externalReferences', CT_ExternalReferences)
register_element_cls('ssml:definedNames', CT_DefinedNames)
register_element_cls('ssml:calcPr', CT_CalcPr)
register_element_cls('ssml:oleSize', CT_OleSize)
register_element_cls('ssml:customWorkbookViews', CT_CustomWorkbookViews)
register_element_cls('ssml:pivotCaches', CT_PivotCaches)
register_element_cls('ssml:smartTagPr', CT_SmartTagPr)
register_element_cls('ssml:smartTagTypes', CT_SmartTagTypes)
register_element_cls('ssml:webPublishing', CT_WebPublishing)
register_element_cls('ssml:fileRecoveryPr', CT_FileRecoveryPr)
register_element_cls('ssml:webPublishObjects', CT_WebPublishObjects)
register_element_cls('ssml:workbookView', CT_BookView)
register_element_cls('ssml:customWorkbookView', CT_CustomWorkbookView)
register_element_cls('ssml:sheet', CT_Sheet)
register_element_cls('ssml:smartTagType', CT_SmartTagType)
register_element_cls('ssml:definedName', CT_DefinedName)
register_element_cls('ssml:externalReference', CT_ExternalReference)
register_element_cls('ssml:pivotCache', CT_PivotCache)
register_element_cls('ssml:functionGroup', CT_FunctionGroup)
register_element_cls('ssml:webPublishObject', CT_WebPublishObject)

# auto_filter
register_element_cls('ssml:filterColumn', CT_FilterColumn)
register_element_cls('ssml:filters', CT_Filters)
register_element_cls('ssml:top10', CT_Top10)
register_element_cls('ssml:customFilters', CT_CustomFilters)
register_element_cls('ssml:dynamicFilter', CT_DynamicFilter)
register_element_cls('ssml:colorFilter', CT_ColorFilter)
register_element_cls('ssml:iconFilter', CT_IconFilter)
register_element_cls('ssml:filter', CT_Filter)
register_element_cls('ssml:dateGroupItem', CT_DateGroupItem)
register_element_cls('ssml:customFilter', CT_CustomFilter)
register_element_cls('ssml:sortCondition', CT_SortCondition)

# pivot_table
register_element_cls('ssml:references', CT_PivotAreaReferences)
register_element_cls('ssml:reference', CT_PivotAreaReference)
register_element_cls('ssml:x', CT_Index)


