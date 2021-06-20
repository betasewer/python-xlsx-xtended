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
    Complex type (sml-styles.xsd)
    
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
    Complex type (sml-styles.xsd)
    
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
    Complex type (sml-styles.xsd)
    
    """
    border = ZeroOrMore('ssml:border') # Border
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Border Count
    
class CT_Border(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
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
    outline = OptionalAttribute('outline', XsdBoolean, True) # Outline
    
class CT_BorderPr(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    color = ZeroOrOne('ssml:color') # Color
    
    style = OptionalAttribute('style', ST_BorderStyle, 'none') # Line Style
    
class CT_CellProtection(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    locked = OptionalAttribute('locked', XsdBoolean) # Cell Locked
    hidden = OptionalAttribute('hidden', XsdBoolean) # Hidden Cell
    
class CT_Fonts(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    font = ZeroOrMore('ssml:font') # Font
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Font Count
    
class CT_Fills(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    fill = ZeroOrMore('ssml:fill') # Fill
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Fill Count
    
class CT_Fill(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    patternFill = ZeroOrOne('ssml:patternFill') # Pattern
    gradientFill = ZeroOrOne('ssml:gradientFill') # Gradient
    
    
class CT_PatternFill(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    fgColor = ZeroOrOne('ssml:fgColor') # Foreground Color
    bgColor = ZeroOrOne('ssml:bgColor') # Background Color
    
    patternType = OptionalAttribute('patternType', ST_PatternType) # Pattern Type
    
class CT_Color(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    auto = OptionalAttribute('auto', XsdBoolean) # Automatic
    indexed = OptionalAttribute('indexed', XsdUnsignedInt) # Index
    rgb = OptionalAttribute('rgb', ST_UnsignedIntHex) # Alpha Red Green Blue Color Value
    theme = OptionalAttribute('theme', XsdUnsignedInt) # Theme Color
    tint = OptionalAttribute('tint', XsdDouble, 0.0) # Tint
    
class CT_GradientFill(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    stop = ZeroOrMore('ssml:stop') # Gradient Stop
    
    type = OptionalAttribute('type', ST_GradientType, 'linear') # Gradient Fill Type
    degree = OptionalAttribute('degree', XsdDouble, 0.0) # Linear Gradient Degree
    left = OptionalAttribute('left', XsdDouble, 0.0) # Left Convergence
    right = OptionalAttribute('right', XsdDouble, 0.0) # Right Convergence
    top = OptionalAttribute('top', XsdDouble, 0.0) # Top Gradient Convergence
    bottom = OptionalAttribute('bottom', XsdDouble, 0.0) # Bottom Convergence
    
class CT_GradientStop(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    color = OneAndOnlyOne('ssml:color') # Color
    
    position = RequiredAttribute('position', XsdDouble) # Gradient Stop Position
    
class CT_NumFmts(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    numFmt = ZeroOrMore('ssml:numFmt') # Number Formats
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Number Format Count
    
class CT_NumFmt(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    numFmtId = RequiredAttribute('numFmtId', ST_NumFmtId) # Number Format Id
    formatCode = RequiredAttribute('formatCode', ST_Xstring) # Number Format Code
    
class CT_CellStyleXfs(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    xf = OneOrMore('ssml:xf') # Formatting Elements
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Style Count
    
class CT_CellXfs(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    xf = OneOrMore('ssml:xf') # Format
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Format Count
    
class CT_Xf(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    alignment = ZeroOrOne('ssml:alignment') # Alignment
    protection = ZeroOrOne('ssml:protection') # Protection
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    numFmtId = OptionalAttribute('numFmtId', ST_NumFmtId) # Number Format Id
    fontId = OptionalAttribute('fontId', ST_FontId) # Font Id
    fillId = OptionalAttribute('fillId', ST_FillId) # Fill Id
    borderId = OptionalAttribute('borderId', ST_BorderId) # Border Id
    xfId = OptionalAttribute('xfId', ST_CellStyleXfId) # Format Id
    quotePrefix = OptionalAttribute('quotePrefix', XsdBoolean, False) # Quote Prefix
    pivotButton = OptionalAttribute('pivotButton', XsdBoolean, False) # Pivot Button
    applyNumberFormat = OptionalAttribute('applyNumberFormat', XsdBoolean) # Apply Number Format
    applyFont = OptionalAttribute('applyFont', XsdBoolean) # Apply Font
    applyFill = OptionalAttribute('applyFill', XsdBoolean) # Apply Fill
    applyBorder = OptionalAttribute('applyBorder', XsdBoolean) # Apply Border
    applyAlignment = OptionalAttribute('applyAlignment', XsdBoolean) # Apply Alignment
    applyProtection = OptionalAttribute('applyProtection', XsdBoolean) # Apply Protection
    
class CT_CellStyles(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    cellStyle = OneOrMore('ssml:cellStyle') # Cell Style
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Style Count
    
class CT_CellStyle(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
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
    Complex type (sml-styles.xsd)
    
    """
    dxf = ZeroOrMore('ssml:dxf') # Formatting
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Format Count
    
class CT_Dxf(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
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
    Complex type (sml-styles.xsd)
    
    """
    indexedColors = ZeroOrOne('ssml:indexedColors') # Color Indexes
    mruColors = ZeroOrOne('ssml:mruColors') # MRU Colors
    
    
class CT_IndexedColors(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    rgbColor = OneOrMore('ssml:rgbColor') # RGB Color
    
    
class CT_MRUColors(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    color = OneOrMore('ssml:color') # Color
    
    
class CT_RgbColor(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    rgb = OptionalAttribute('rgb', ST_UnsignedIntHex) # Alpha Red Green Blue
    
class CT_TableStyles(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    tableStyle = ZeroOrMore('ssml:tableStyle') # Table Style
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Table Style Count
    defaultTableStyle = OptionalAttribute('defaultTableStyle', XsdString) # Default Table Style
    defaultPivotStyle = OptionalAttribute('defaultPivotStyle', XsdString) # Default Pivot Style
    
class CT_TableStyle(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    tableStyleElement = ZeroOrMore('ssml:tableStyleElement') # Table Style
    
    name = RequiredAttribute('name', XsdString) # Table Style Name
    pivot = OptionalAttribute('pivot', XsdBoolean, True) # Pivot Style
    table = OptionalAttribute('table', XsdBoolean, True) # Table
    count = OptionalAttribute('count', XsdUnsignedInt) # Table Style Count
    
class CT_TableStyleElement(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    type = RequiredAttribute('type', ST_TableStyleType) # Table Style Type
    size = OptionalAttribute('size', XsdUnsignedInt, 1) # Band Size
    dxfId = OptionalAttribute('dxfId', ST_DxfId) # Formatting Id
    
class CT_BooleanProperty(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    val = OptionalAttribute('val', XsdBoolean, True) # Value
    
class CT_FontSize(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    val = RequiredAttribute('val', XsdDouble) # Value
    
class CT_IntProperty(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    val = RequiredAttribute('val', XsdInt) # Value
    
class CT_FontName(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    val = RequiredAttribute('val', ST_Xstring) # String Value
    
class CT_VerticalAlignFontProperty(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    val = RequiredAttribute('val', ST_VerticalAlignRun) # Value
    
class CT_FontScheme(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    val = RequiredAttribute('val', ST_FontScheme) # Font Scheme
    
class CT_UnderlineProperty(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
    """
    val = OptionalAttribute('val', ST_UnderlineValues, 'single') # Underline Value
    
class CT_Font(BaseOxmlElement):
    """
    Complex type (sml-styles.xsd)
    
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
   
   
