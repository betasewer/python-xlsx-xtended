# encoding: utf-8

"""
Custom element classes that correspond to the document part, e.g.
<w:document>.
"""

from docxx.oxml.xmlchemy import (
    BaseOxmlElement, ZeroOrOne, ZeroOrMore, OneOrMore, OneAndOnlyOne,
    RequiredAttribute, OptionalAttribute, 
)
from docxx.oxml.simpletypes import (
    XsdUnsignedInt, ST_RelationshipId, XsdDouble, XsdBoolean, XsdUnsignedByte
)
from xlsxx.oxml.simpletypes import (
    ST_Xstring, ST_Ref, ST_CellRef, ST_CellType, ST_CellSpans
)

#
#
#
#
#
#
class CT_Worksheet(BaseOxmlElement):
    dimension = ZeroOrOne('ssml:dimension')
    cols = ZeroOrMore('ssml:cols')
    sheetdata = OneAndOnlyOne('ssml:sheetData')
    
class CT_SheetDimension(BaseOxmlElement):
    ref = RequiredAttribute('ref', ST_Ref) # Reference
    
class CT_Cols(BaseOxmlElement):
    col = OneOrMore('ssml:col')
    
class CT_Col(BaseOxmlElement):
    min = RequiredAttribute("min", XsdUnsignedInt)	
    max = RequiredAttribute("max", XsdUnsignedInt)
    width = OptionalAttribute("width", XsdDouble)
    style = OptionalAttribute("style", XsdUnsignedInt)

class CT_SheetData(BaseOxmlElement):
    row = ZeroOrMore('ssml:row')

class CT_Row(BaseOxmlElement):
    cell = ZeroOrMore('ssml:c') # Cell
    #extLst = *ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    ref = OptionalAttribute('r', XsdUnsignedInt) # Row Index
    spans = OptionalAttribute('spans', ST_CellSpans) # Spans
    style = OptionalAttribute('s', XsdUnsignedInt) # Style Index
    customFormat = OptionalAttribute('customFormat', XsdBoolean) # Custom Format
    ht = OptionalAttribute('ht', XsdDouble) # Row Height
    hidden = OptionalAttribute('hidden', XsdBoolean) # Hidden
    customHeight = OptionalAttribute('customHeight', XsdBoolean) # Custom Height
    outlineLevel = OptionalAttribute('outlineLevel', XsdUnsignedByte) # Outline Level
    collapsed = OptionalAttribute('collapsed', XsdBoolean) # Collapsed
    thickTop = OptionalAttribute('thickTop', XsdBoolean) # Thick Top Border
    thickBot = OptionalAttribute('thickBot', XsdBoolean) # Thick Bottom
    phonetic = OptionalAttribute('ph', XsdBoolean) # Show Phonetic
    
    
class CT_Cell(BaseOxmlElement):
    formula = ZeroOrOne('ssml:f') # Formula
    value = ZeroOrOne('ssml:v') # Cell Value
    inline_text = ZeroOrOne('ssml:is') # Rich Text Inline
    #extLst = *ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    ref = OptionalAttribute('r', ST_CellRef) # Reference	
    style = OptionalAttribute('s', XsdUnsignedInt) # Style Index Default value is "0".
    type = OptionalAttribute('t', ST_CellType) #	Cell Data Type	Default value is "n".
    cell_metadata = OptionalAttribute('cm', XsdUnsignedInt) #	Cell Metadata Index	Default value is "0".
    value_metadata = OptionalAttribute('vm', XsdUnsignedInt) #	Value Metadata Index	Default value is "0".
    phonetic = OptionalAttribute('ph', XsdBoolean) #	Show Phonetic	Default value is "false".    
    
class CT_CellFormula(BaseOxmlElement):
    """
    Complex type: 
    """
    # simpleContent[]: 
    
#
# requried: ST_Ref, XsdUnsignedInt
#
class CT_MergeCells(BaseOxmlElement):
    """
    Complex type: 
    """
    mergeCell = OneOrMore('ssml:mergeCell') # Merged Cell
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Count
    
class CT_MergeCell(BaseOxmlElement):
    """
    Complex type: 
    """
    ref = RequiredAttribute('ref', ST_Ref) # Reference
    
