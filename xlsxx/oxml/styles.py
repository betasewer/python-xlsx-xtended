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
    ST_VerticalAlignRun, XsdBoolean, XsdDouble, XsdInt, XsdString, XsdUnsignedInt,
)
from xlsxx.oxml.simpletypes import (
    ST_FontId, ST_UnsignedIntHex, ST_Xstring,
    ST_BorderId, ST_BorderStyle, ST_CellStyleXfId, ST_DxfId, ST_FillId, ST_FontScheme, 
    ST_GradientType, ST_HorizontalAlignment, ST_NumFmtId, ST_PatternType, ST_TableStyleType, 
    ST_UnderlineValues, ST_VerticalAlignment,

)

class CT_Stylesheet(BaseOxmlElement):
    """
    Complex type: 
    """
    numFmts = ZeroOrOne('ssml:numFmts') # Number Formats
    fonts = ZeroOrOne('ssml:fonts') # Fonts
    fills = ZeroOrOne('ssml:fills') # Fills
    borders = ZeroOrOne('ssml:borders') # Borders
    cellStyleXfs = ZeroOrOne('ssml:cellStyleXfs') # Formatting Records
    cellXfs = ZeroOrOne('ssml:cellXfs') # Cell Formats
    cellStyles = ZeroOrOne('ssml:cellStyles') # Cell Styles
    dxfs = ZeroOrOne('ssml:dxfs') # Formats
    tableStyles = ZeroOrOne('ssml:tableStyles') # Table Styles
    colors = ZeroOrOne('ssml:colors') # Colors
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    
class CT_CellAlignment(BaseOxmlElement):
    """
    Complex type: 
    """
    horizontal = OptionalAttribute('horizontal', ST_HorizontalAlignment) # Horizontal Alignment
    vertical = OptionalAttribute('vertical', ST_VerticalAlignment) # Vertical Alignment
    textRotation = OptionalAttribute('textRotation', XsdUnsignedInt) # Text Rotation
    wrapText = OptionalAttribute('wrapText', XsdBoolean) # Wrap Text
    indent = OptionalAttribute('indent', XsdUnsignedInt) # Indent
    relativeIndent = OptionalAttribute('relativeIndent', XsdInt) # Relative Indent
    justifyLastLine = OptionalAttribute('justifyLastLine', XsdBoolean) # Justify Last Line
    shrinkToFit = OptionalAttribute('shrinkToFit', XsdBoolean) # Shrink To Fit
    readingOrder = OptionalAttribute('readingOrder', XsdUnsignedInt) # Reading Order
    
class CT_Borders(BaseOxmlElement):
    """
    Complex type: 
    """
    border = ZeroOrMore('ssml:border') # Border
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Border Count
    
class CT_Border(BaseOxmlElement):
    """
    Complex type: 
    """
    left = ZeroOrOne('ssml:left') # Left Border
    right = ZeroOrOne('ssml:right') # Right Border
    top = ZeroOrOne('ssml:top') # Top Border
    bottom = ZeroOrOne('ssml:bottom') # Bottom Border
    diagonal = ZeroOrOne('ssml:diagonal') # Diagonal
    vertical = ZeroOrOne('ssml:vertical') # Vertical Inner Border
    horizontal = ZeroOrOne('ssml:horizontal') # Horizontal Inner Borders
    
    diagonalUp = OptionalAttribute('diagonalUp', XsdBoolean) # Diagonal Up
    diagonalDown = OptionalAttribute('diagonalDown', XsdBoolean) # Diagonal Down
    outline = OptionalAttribute('outline', XsdBoolean) # Outline
    
class CT_BorderPr(BaseOxmlElement):
    """
    Complex type: 
    """
    color = ZeroOrOne('ssml:color') # Color
    
    style = OptionalAttribute('style', ST_BorderStyle) # Line Style
    
class CT_CellProtection(BaseOxmlElement):
    """
    Complex type: 
    """
    locked = OptionalAttribute('locked', XsdBoolean) # Cell Locked
    hidden = OptionalAttribute('hidden', XsdBoolean) # Hidden Cell
    
class CT_Fonts(BaseOxmlElement):
    """
    Complex type: 
    """
    font = ZeroOrMore('ssml:font') # Font
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Font Count
    
class CT_Fills(BaseOxmlElement):
    """
    Complex type: 
    """
    fill = ZeroOrMore('ssml:fill') # Fill
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Fill Count
    
class CT_Fill(BaseOxmlElement):
    """
    Complex type: 
    """
    patternFill = ZeroOrOne('ssml:patternFill')
    gradientFill = ZeroOrOne('ssml:gradientFill')
      
    
class CT_PatternFill(BaseOxmlElement):
    """
    Complex type: 
    """
    fgColor = ZeroOrOne('ssml:fgColor') # Foreground Color
    bgColor = ZeroOrOne('ssml:bgColor') # Background Color
    
    patternType = OptionalAttribute('patternType', ST_PatternType) # Pattern Type
    
class CT_Color(BaseOxmlElement):
    """
    Complex type: 
    """
    auto = OptionalAttribute('auto', XsdBoolean) # Automatic
    indexed = OptionalAttribute('indexed', XsdUnsignedInt) # Index
    rgb = OptionalAttribute('rgb', ST_UnsignedIntHex) # Alpha Red Green Blue Color Value
    theme = OptionalAttribute('theme', XsdUnsignedInt) # Theme Color
    tint = OptionalAttribute('tint', XsdDouble) # Tint
    
class CT_GradientFill(BaseOxmlElement):
    """
    Complex type: 
    """
    stop = ZeroOrMore('ssml:stop') # Gradient Stop
    
    type = OptionalAttribute('type', ST_GradientType) # Gradient Fill Type
    degree = OptionalAttribute('degree', XsdDouble) # Linear Gradient Degree
    left = OptionalAttribute('left', XsdDouble) # Left Convergence
    right = OptionalAttribute('right', XsdDouble) # Right Convergence
    top = OptionalAttribute('top', XsdDouble) # Top Gradient Convergence
    bottom = OptionalAttribute('bottom', XsdDouble) # Bottom Convergence
    
class CT_GradientStop(BaseOxmlElement):
    """
    Complex type: 
    """
    color = OneAndOnlyOne('ssml:color') # Color
    
    position = RequiredAttribute('position', XsdDouble) # Gradient Stop Position
    
class CT_NumFmts(BaseOxmlElement):
    """
    Complex type: 
    """
    numFmt = ZeroOrMore('ssml:numFmt') # Number Formats
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Number Format Count
    
class CT_NumFmt(BaseOxmlElement):
    """
    Complex type: 
    """
    numFmtId = RequiredAttribute('numFmtId', ST_NumFmtId) # Number Format Id
    formatCode = RequiredAttribute('formatCode', ST_Xstring) # Number Format Code
    
class CT_CellStyleXfs(BaseOxmlElement):
    """
    Complex type: 
    """
    xf = OneOrMore('ssml:xf') # Formatting Elements
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Style Count
    
class CT_CellXfs(BaseOxmlElement):
    """
    Complex type: 
    """
    xf = OneOrMore('ssml:xf') # Format
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Format Count
    
class CT_Xf(BaseOxmlElement):
    """
    Complex type: 
    """
    alignment = ZeroOrOne('ssml:alignment') # Alignment
    protection = ZeroOrOne('ssml:protection') # Protection
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    numFmtId = OptionalAttribute('numFmtId', ST_NumFmtId) # Number Format Id
    fontId = OptionalAttribute('fontId', ST_FontId) # Font Id
    fillId = OptionalAttribute('fillId', ST_FillId) # Fill Id
    borderId = OptionalAttribute('borderId', ST_BorderId) # Border Id
    xfId = OptionalAttribute('xfId', ST_CellStyleXfId) # Format Id
    quotePrefix = OptionalAttribute('quotePrefix', XsdBoolean) # Quote Prefix
    pivotButton = OptionalAttribute('pivotButton', XsdBoolean) # Pivot Button
    applyNumberFormat = OptionalAttribute('applyNumberFormat', XsdBoolean) # Apply Number Format
    applyFont = OptionalAttribute('applyFont', XsdBoolean) # Apply Font
    applyFill = OptionalAttribute('applyFill', XsdBoolean) # Apply Fill
    applyBorder = OptionalAttribute('applyBorder', XsdBoolean) # Apply Border
    applyAlignment = OptionalAttribute('applyAlignment', XsdBoolean) # Apply Alignment
    applyProtection = OptionalAttribute('applyProtection', XsdBoolean) # Apply Protection
    
class CT_CellStyles(BaseOxmlElement):
    """
    Complex type: 
    """
    cellStyle = OneOrMore('ssml:cellStyle') # Cell Style
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Style Count
    
class CT_CellStyle(BaseOxmlElement):
    """
    Complex type: 
    """
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    name = OptionalAttribute('name', ST_Xstring) # User Defined Cell Style
    xfId = RequiredAttribute('xfId', ST_CellStyleXfId) # Format Id
    builtinId = OptionalAttribute('builtinId', XsdUnsignedInt) # Built-In Style Id
    iLevel = OptionalAttribute('iLevel', XsdUnsignedInt) # Outline Style
    hidden = OptionalAttribute('hidden', XsdBoolean) # Hidden Style
    customBuiltin = OptionalAttribute('customBuiltin', XsdBoolean) # Custom Built In
    
class CT_Dxfs(BaseOxmlElement):
    """
    Complex type: 
    """
    dxf = ZeroOrMore('ssml:dxf') # Formatting
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Format Count
    
class CT_Dxf(BaseOxmlElement):
    """
    Complex type: 
    """
    font = ZeroOrOne('ssml:font') # Font Properties
    numFmt = ZeroOrOne('ssml:numFmt') # Number Format
    fill = ZeroOrOne('ssml:fill') # Fill
    alignment = ZeroOrOne('ssml:alignment') # Alignment
    border = ZeroOrOne('ssml:border') # Border Properties
    protection = ZeroOrOne('ssml:protection') # Protection Properties
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    
class CT_Colors(BaseOxmlElement):
    """
    Complex type: 
    """
    indexedColors = ZeroOrOne('ssml:indexedColors') # Color Indexes
    mruColors = ZeroOrOne('ssml:mruColors') # MRU Colors
    
    
class CT_IndexedColors(BaseOxmlElement):
    """
    Complex type: 
    """
    rgbColor = OneOrMore('ssml:rgbColor') # RGB Color
    
    
class CT_MRUColors(BaseOxmlElement):
    """
    Complex type: 
    """
    color = OneOrMore('ssml:color') # Color
    
    
class CT_RgbColor(BaseOxmlElement):
    """
    Complex type: 
    """
    rgb = OptionalAttribute('rgb', ST_UnsignedIntHex) # Alpha Red Green Blue
    
class CT_TableStyles(BaseOxmlElement):
    """
    Complex type: 
    """
    tableStyle = ZeroOrMore('ssml:tableStyle') # Table Style
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Table Style Count
    defaultTableStyle = OptionalAttribute('defaultTableStyle', XsdString) # Default Table Style
    defaultPivotStyle = OptionalAttribute('defaultPivotStyle', XsdString) # Default Pivot Style
    
class CT_TableStyle(BaseOxmlElement):
    """
    Complex type: 
    """
    tableStyleElement = ZeroOrMore('ssml:tableStyleElement') # Table Style
    
    name = RequiredAttribute('name', XsdString) # Table Style Name
    pivot = OptionalAttribute('pivot', XsdBoolean) # Pivot Style
    table = OptionalAttribute('table', XsdBoolean) # Table
    count = OptionalAttribute('count', XsdUnsignedInt) # Table Style Count
    
class CT_TableStyleElement(BaseOxmlElement):
    """
    Complex type: 
    """
    type = RequiredAttribute('type', ST_TableStyleType) # Table Style Type
    size = OptionalAttribute('size', XsdUnsignedInt) # Band Size
    dxfId = OptionalAttribute('dxfId', ST_DxfId) # Formatting Id
    
class CT_BooleanProperty(BaseOxmlElement):
    """
    Complex type: 
    """
    val = OptionalAttribute('val', XsdBoolean) # Value
    
class CT_FontSize(BaseOxmlElement):
    """
    Complex type: 
    """
    val = RequiredAttribute('val', XsdDouble) # Value
    
class CT_IntProperty(BaseOxmlElement):
    """
    Complex type: 
    """
    val = RequiredAttribute('val', XsdInt) # Value
    
class CT_FontName(BaseOxmlElement):
    """
    Complex type: 
    """
    val = RequiredAttribute('val', ST_Xstring) # String Value
    
class CT_VerticalAlignFontProperty(BaseOxmlElement):
    """
    Complex type: 
    """
    val = RequiredAttribute('val', ST_VerticalAlignRun) # Value
    
class CT_FontScheme(BaseOxmlElement):
    """
    Complex type: 
    """
    val = RequiredAttribute('val', ST_FontScheme) # Font Scheme
    
class CT_UnderlineProperty(BaseOxmlElement):
    """
    Complex type: 
    """
    val = OptionalAttribute('val', ST_UnderlineValues) # Underline Value
    
class CT_Font(BaseOxmlElement):
    """
    Complex type: 
    """
    name = ZeroOrOne('ssml:name') # Font Name
    charset = ZeroOrOne('ssml:charset') # Character Set
    family = ZeroOrOne('ssml:family') # Font Family
    b = ZeroOrOne('ssml:b') # Bold
    i = ZeroOrOne('ssml:i') # Italic
    strike = ZeroOrOne('ssml:strike') # Strike Through
    outline = ZeroOrOne('ssml:outline') # Outline
    shadow = ZeroOrOne('ssml:shadow') # Shadow
    condense = ZeroOrOne('ssml:condense') # Condense
    extend = ZeroOrOne('ssml:extend') # Extend
    color = ZeroOrOne('ssml:color') # Text Color
    sz = ZeroOrOne('ssml:sz') # Font Size
    u = ZeroOrOne('ssml:u') # Underline
    vertAlign = ZeroOrOne('ssml:vertAlign') # Text Vertical Alignment
    scheme = ZeroOrOne('ssml:scheme') # Scheme
