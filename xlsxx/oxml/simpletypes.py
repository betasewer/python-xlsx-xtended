# encoding: utf-8

"""
Custom element classes that correspond to the document part, e.g.
<w:document>.
"""


from docxx.oxml.simpletypes import (
    XsdString, XsdStringEnumeration, XsdUnsignedInt, ST_RelationshipId, XsdToken
)

#
#
# sml
#
#
class ST_Xstring(XsdString):
    pass
    
class ST_Ref(XsdString):
    pass
    
class ST_CellRef(XsdString):
    pass

class ST_RefA(XsdString):
    pass
    
class ST_Sqref():
    """
    Simple type: Reference Sequence
    """
    # list[itemType='ST_Ref']: None
    
class ST_Formula(ST_Xstring):
    pass
    
class ST_UnsignedIntHex(XsdUnsignedInt): # xsd:hexBinary
    """
    Simple type: Hex Unsigned Integer
    """
    # length[value='4']: None
    
class ST_UnsignedShortHex(XsdUnsignedInt): # xsd:hexBinary
    """
    Simple type: Unsigned Short Hex
    """
    # length[value='2']: None
    
class ST_Guid(XsdToken):
    """
    Simple type: Globally Unique Identifier
    """
    # pattern[value='\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}']: None

#
# sheet.xsd
#
class ST_CellType(XsdStringEnumeration):
    boolean = "b"
    number = "n"
    error = "e"
    shared_string = "s"
    string = "str"
    inline_string = "inlineStr"
    
    _members = (boolean,number,error,shared_string,string,inline_string)

class ST_CellSpans(XsdString):
    # list[itemType='ST_CellSpan']: None
    pass

#
# sharedstringtable.xsd
#
class ST_PhoneticType(XsdStringEnumeration):
    """
    Simple type: Phonetic Type
    """
    HALFWIDTH_KATAKANA = 'halfwidthKatakana'  # Half-Width Katakana
    FULLWIDTH_KATAKANA = 'fullwidthKatakana'  # Full-Width Katakana
    HIRAGANA = 'Hiragana'  # Hiragana
    NO_CONVERSION = 'noConversion'  # No Conversion
    _members = (HALFWIDTH_KATAKANA, FULLWIDTH_KATAKANA, HIRAGANA, NO_CONVERSION,)
    
class ST_PhoneticAlignment(XsdStringEnumeration):
    """
    Simple type: Phonetic Alignment Types
    """
    NO_CONTROL = 'noControl'  # No Control
    LEFT = 'left'  # Left Alignment
    CENTER = 'center'  # Center Alignment
    DISTRIBUTED = 'distributed'  # Distributed
    _members = (NO_CONTROL, LEFT, CENTER, DISTRIBUTED,)
    
#
# styles.xsd
#
class ST_BorderStyle(XsdStringEnumeration):
    """
    Simple type: Border Line Styles
    """
    NONE = 'none'  # None
    THIN = 'thin'  # Thin Border
    MEDIUM = 'medium'  # Medium Border
    DASHED = 'dashed'  # Dashed
    DOTTED = 'dotted'  # Dotted
    THICK = 'thick'  # Thick Line Border
    DOUBLE = 'double'  # Double Line
    HAIR = 'hair'  # Hairline Border
    MEDIUM_DASHED = 'mediumDashed'  # Medium Dashed
    DASH_DOT = 'dashDot'  # Dash Dot
    MEDIUM_DASH_DOT = 'mediumDashDot'  # Medium Dash Dot
    DASH_DOT_DOT = 'dashDotDot'  # Dash Dot Dot
    MEDIUM_DASH_DOT_DOT = 'mediumDashDotDot'  # Medium Dash Dot Dot
    SLANT_DASH_DOT = 'slantDashDot'  # Slant Dash Dot
    _members = (NONE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, DASH_DOT, MEDIUM_DASH_DOT, DASH_DOT_DOT, MEDIUM_DASH_DOT_DOT, SLANT_DASH_DOT,)
    
class ST_PatternType(XsdStringEnumeration):
    """
    Simple type: Pattern Type
    """
    NONE = 'none'  # None
    SOLID = 'solid'  # Solid
    MEDIUM_GRAY = 'mediumGray'  # Medium Gray
    DARK_GRAY = 'darkGray'  # Dary Gray
    LIGHT_GRAY = 'lightGray'  # Light Gray
    DARK_HORIZONTAL = 'darkHorizontal'  # Dark Horizontal
    DARK_VERTICAL = 'darkVertical'  # Dark Vertical
    DARK_DOWN = 'darkDown'  # Dark Down
    DARK_UP = 'darkUp'  # Dark Up
    DARK_GRID = 'darkGrid'  # Dark Grid
    DARK_TRELLIS = 'darkTrellis'  # Dark Trellis
    LIGHT_HORIZONTAL = 'lightHorizontal'  # Light Horizontal
    LIGHT_VERTICAL = 'lightVertical'  # Light Vertical
    LIGHT_DOWN = 'lightDown'  # Light Down
    LIGHT_UP = 'lightUp'  # Light Up
    LIGHT_GRID = 'lightGrid'  # Light Grid
    LIGHT_TRELLIS = 'lightTrellis'  # Light Trellis
    GRAY125 = 'gray125'  # Gray 0.125
    GRAY0625 = 'gray0625'  # Gray 0.0625
    _members = (NONE, SOLID, MEDIUM_GRAY, DARK_GRAY, LIGHT_GRAY, DARK_HORIZONTAL, DARK_VERTICAL, DARK_DOWN, DARK_UP, DARK_GRID, DARK_TRELLIS, LIGHT_HORIZONTAL, LIGHT_VERTICAL, LIGHT_DOWN, LIGHT_UP, LIGHT_GRID, LIGHT_TRELLIS, GRAY125, GRAY0625,)
    
class ST_GradientType(XsdStringEnumeration):
    """
    Simple type: Gradient Type
    """
    LINEAR = 'linear'  # Linear Gradient
    PATH = 'path'  # Path
    _members = (LINEAR, PATH,)
    
class ST_HorizontalAlignment(XsdStringEnumeration):
    """
    Simple type: Horizontal Alignment Type
    """
    GENERAL = 'general'  # General Horizontal Alignment
    LEFT = 'left'  # Left Horizontal Alignment
    CENTER = 'center'  # Centered Horizontal Alignment
    RIGHT = 'right'  # Right Horizontal Alignment
    FILL = 'fill'  # Fill
    JUSTIFY = 'justify'  # Justify
    CENTER_CONTINUOUS = 'centerContinuous'  # Center Continuous Horizontal Alignment
    DISTRIBUTED = 'distributed'  # Distributed Horizontal Alignment
    _members = (GENERAL, LEFT, CENTER, RIGHT, FILL, JUSTIFY, CENTER_CONTINUOUS, DISTRIBUTED,)
    
class ST_VerticalAlignment(XsdStringEnumeration):
    """
    Simple type: Vertical Alignment Types
    """
    TOP = 'top'  # Align Top
    CENTER = 'center'  # Centered Vertical Alignment
    BOTTOM = 'bottom'  # Aligned To Bottom
    JUSTIFY = 'justify'  # Justified Vertically
    DISTRIBUTED = 'distributed'  # Distributed Vertical Alignment
    _members = (TOP, CENTER, BOTTOM, JUSTIFY, DISTRIBUTED,)
    
class ST_NumFmtId(XsdUnsignedInt):
    """
    Simple type: Number Format Id
    """
    
class ST_FontId(XsdUnsignedInt):
    """
    Simple type: Font Id
    """
    
class ST_FillId(XsdUnsignedInt):
    """
    Simple type: Fill Id
    """
    
class ST_BorderId(XsdUnsignedInt):
    """
    Simple type: Border Id
    """
    
class ST_CellStyleXfId(XsdUnsignedInt):
    """
    Simple type: Cell Style Format Id
    """
    
class ST_DxfId(XsdUnsignedInt):
    """
    Simple type: Format Id
    """
    
class ST_TableStyleType(XsdStringEnumeration):
    """
    Simple type: Table Style Type
    """
    WHOLE_TABLE = 'wholeTable'  # Whole Table Style
    HEADER_ROW = 'headerRow'  # Header Row Style
    TOTAL_ROW = 'totalRow'  # Total Row Style
    FIRST_COLUMN = 'firstColumn'  # First Column Style
    LAST_COLUMN = 'lastColumn'  # Last Column Style
    FIRST_ROW_STRIPE = 'firstRowStripe'  # First Row Stripe Style
    SECOND_ROW_STRIPE = 'secondRowStripe'  # Second Row Stripe Style
    FIRST_COLUMN_STRIPE = 'firstColumnStripe'  # First Column Stripe Style
    SECOND_COLUMN_STRIPE = 'secondColumnStripe'  # Second Column Stipe Style
    FIRST_HEADER_CELL = 'firstHeaderCell'  # First Header Row Style
    LAST_HEADER_CELL = 'lastHeaderCell'  # Last Header Style
    FIRST_TOTAL_CELL = 'firstTotalCell'  # First Total Row Style
    LAST_TOTAL_CELL = 'lastTotalCell'  # Last Total Row Style
    FIRST_SUBTOTAL_COLUMN = 'firstSubtotalColumn'  # First Subtotal Column Style
    SECOND_SUBTOTAL_COLUMN = 'secondSubtotalColumn'  # Second Subtotal Column Style
    THIRD_SUBTOTAL_COLUMN = 'thirdSubtotalColumn'  # Third Subtotal Column Style
    FIRST_SUBTOTAL_ROW = 'firstSubtotalRow'  # First Subtotal Row Style
    SECOND_SUBTOTAL_ROW = 'secondSubtotalRow'  # Second Subtotal Row Style
    THIRD_SUBTOTAL_ROW = 'thirdSubtotalRow'  # Third Subtotal Row Style
    BLANK_ROW = 'blankRow'  # Blank Row Style
    FIRST_COLUMN_SUBHEADING = 'firstColumnSubheading'  # First Column Subheading Style
    SECOND_COLUMN_SUBHEADING = 'secondColumnSubheading'  # Second Column Subheading Style
    THIRD_COLUMN_SUBHEADING = 'thirdColumnSubheading'  # Third Column Subheading Style
    FIRST_ROW_SUBHEADING = 'firstRowSubheading'  # First Row Subheading Style
    SECOND_ROW_SUBHEADING = 'secondRowSubheading'  # Second Row Subheading Style
    THIRD_ROW_SUBHEADING = 'thirdRowSubheading'  # Third Row Subheading Style
    PAGE_FIELD_LABELS = 'pageFieldLabels'  # Page Field Labels Style
    PAGE_FIELD_VALUES = 'pageFieldValues'  # Page Field Values Style
    _members = (WHOLE_TABLE, HEADER_ROW, TOTAL_ROW, FIRST_COLUMN, LAST_COLUMN, FIRST_ROW_STRIPE, SECOND_ROW_STRIPE, FIRST_COLUMN_STRIPE, SECOND_COLUMN_STRIPE, FIRST_HEADER_CELL, LAST_HEADER_CELL, FIRST_TOTAL_CELL, LAST_TOTAL_CELL, FIRST_SUBTOTAL_COLUMN, SECOND_SUBTOTAL_COLUMN, THIRD_SUBTOTAL_COLUMN, FIRST_SUBTOTAL_ROW, SECOND_SUBTOTAL_ROW, THIRD_SUBTOTAL_ROW, BLANK_ROW, FIRST_COLUMN_SUBHEADING, SECOND_COLUMN_SUBHEADING, THIRD_COLUMN_SUBHEADING, FIRST_ROW_SUBHEADING, SECOND_ROW_SUBHEADING, THIRD_ROW_SUBHEADING, PAGE_FIELD_LABELS, PAGE_FIELD_VALUES,)
    
class ST_VerticalAlignRun(XsdStringEnumeration):
    """
    Simple type: Vertical Alignment Run Types
    """
    BASELINE = 'baseline'  # Baseline
    SUPERSCRIPT = 'superscript'  # Superscript
    SUBSCRIPT = 'subscript'  # Subscript
    _members = (BASELINE, SUPERSCRIPT, SUBSCRIPT,)
    
class ST_FontScheme(XsdStringEnumeration):
    """
    Simple type: Font scheme Styles
    """
    NONE = 'none'  # None
    MAJOR = 'major'  # Major Font
    MINOR = 'minor'  # Minor Font
    _members = (NONE, MAJOR, MINOR,)
    
class ST_UnderlineValues(XsdStringEnumeration):
    """
    Simple type: Underline Types
    """
    SINGLE = 'single'  # Single Underline
    DOUBLE = 'double'  # Double Underline
    SINGLE_ACCOUNTING = 'singleAccounting'  # Accounting Single Underline
    DOUBLE_ACCOUNTING = 'doubleAccounting'  # Accounting Double Underline
    NONE = 'none'  # None
    _members = (SINGLE, DOUBLE, SINGLE_ACCOUNTING, DOUBLE_ACCOUNTING, NONE,)
    



