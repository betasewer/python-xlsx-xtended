# encoding: utf-8

"""
Custom element classes that correspond to the document part, e.g.
<w:document>.
"""


from docxx.oxml.simpletypes import (
    XsdString, XsdStringEnumeration, 
    XsdUnsignedInt, XsdToken,
    XsdHexBinary
)

#
#
# sml
#
#
class ST_CellSpan(XsdString):
    """
    Simple type: Cell Span Type
    """
    
class ST_CellSpans():
    """
    Simple type: Cell Spans
    """
    # list[itemType='ST_CellSpan']: None
    
class ST_CellType(XsdStringEnumeration):
    """
    Simple type: Cell Type
    """
    BOOLEAN = 'b'  # Boolean
    NUMBER = 'n'  # Number
    ERROR = 'e'  # Error
    SHARED_STRING = 's'  # Shared String
    STR = 'str'  # String
    INLINE_STR = 'inlineStr'  # Inline String
    _members = (BOOLEAN, NUMBER, ERROR, SHARED_STRING, STR, INLINE_STR,)
    
class ST_CellFormulaType(XsdStringEnumeration):
    """
    Simple type: Formula Type
    """
    NORMAL = 'normal'  # Normal
    ARRAY = 'array'  # Array Entered
    DATA_TABLE = 'dataTable'  # Table Formula
    SHARED = 'shared'  # Shared Formula
    _members = (NORMAL, ARRAY, DATA_TABLE, SHARED,)
    
class ST_Pane(XsdStringEnumeration):
    """
    Simple type: Pane Types
    """
    BOTTOM_RIGHT = 'bottomRight'  # Bottom Right Pane
    TOP_RIGHT = 'topRight'  # Top Right Pane
    BOTTOM_LEFT = 'bottomLeft'  # Bottom Left Pane
    TOP_LEFT = 'topLeft'  # Top Left Pane
    _members = (BOTTOM_RIGHT, TOP_RIGHT, BOTTOM_LEFT, TOP_LEFT,)
    
class ST_SheetViewType(XsdStringEnumeration):
    """
    Simple type: Sheet View Type
    """
    NORMAL = 'normal'  # Normal View
    PAGE_BREAK_PREVIEW = 'pageBreakPreview'  # Page Break Preview
    PAGE_LAYOUT = 'pageLayout'  # Page Layout View
    _members = (NORMAL, PAGE_BREAK_PREVIEW, PAGE_LAYOUT,)
    
class ST_DataConsolidateFunction(XsdStringEnumeration):
    """
    Simple type: Data Consolidation Functions
    """
    AVERAGE = 'average'  # Average
    COUNT = 'count'  # Count
    COUNT_NUMS = 'countNums'  # CountNums
    MAX = 'max'  # Maximum
    MIN = 'min'  # Minimum
    PRODUCT = 'product'  # Product
    STD_DEV = 'stdDev'  # StdDev
    STD_DEVP = 'stdDevp'  # StdDevP
    SUM = 'sum'  # Sum
    VAR = 'var'  # Variance
    VARP = 'varp'  # VarP
    _members = (AVERAGE, COUNT, COUNT_NUMS, MAX, MIN, PRODUCT, STD_DEV, STD_DEVP, SUM, VAR, VARP,)
    
class ST_DataValidationType(XsdStringEnumeration):
    """
    Simple type: Data Validation Type
    """
    NONE = 'none'  # None
    WHOLE = 'whole'  # Whole Number
    DECIMAL = 'decimal'  # Decimal
    LIST = 'list'  # List
    DATE = 'date'  # Date
    TIME = 'time'  # Time
    TEXT_LENGTH = 'textLength'  # Text Length
    CUSTOM = 'custom'  # Custom
    _members = (NONE, WHOLE, DECIMAL, LIST, DATE, TIME, TEXT_LENGTH, CUSTOM,)
    
class ST_DataValidationOperator(XsdStringEnumeration):
    """
    Simple type: Data Validation Operator
    """
    BETWEEN = 'between'  # Between
    NOT_BETWEEN = 'notBetween'  # Not Between
    EQUAL = 'equal'  # Equal
    NOT_EQUAL = 'notEqual'  # Not Equal
    LESS_THAN = 'lessThan'  # Less Than
    LESS_THAN_OR_EQUAL = 'lessThanOrEqual'  # Less Than Or Equal
    GREATER_THAN = 'greaterThan'  # Greater Than
    GREATER_THAN_OR_EQUAL = 'greaterThanOrEqual'  # Greater Than Or Equal
    _members = (BETWEEN, NOT_BETWEEN, EQUAL, NOT_EQUAL, LESS_THAN, LESS_THAN_OR_EQUAL, GREATER_THAN, GREATER_THAN_OR_EQUAL,)
    
class ST_DataValidationErrorStyle(XsdStringEnumeration):
    """
    Simple type: Data Validation Error Styles
    """
    STOP = 'stop'  # Stop Icon
    WARNING = 'warning'  # Warning Icon
    INFORMATION = 'information'  # Information Icon
    _members = (STOP, WARNING, INFORMATION,)
    
class ST_DataValidationImeMode(XsdStringEnumeration):
    """
    Simple type: Data Validation IME Mode
    """
    NO_CONTROL = 'noControl'  # IME Mode Not Controlled
    OFF = 'off'  # IME Off
    IME_ON = 'on'  # IME On
    DISABLED = 'disabled'  # Disabled IME Mode
    HIRAGANA = 'hiragana'  # Hiragana IME Mode
    FULL_KATAKANA = 'fullKatakana'  # Full Katakana IME Mode
    HALF_KATAKANA = 'halfKatakana'  # Half-Width Katakana
    FULL_ALPHA = 'fullAlpha'  # Full-Width Alpha-Numeric IME Mode
    HALF_ALPHA = 'halfAlpha'  # Half Alpha IME
    FULL_HANGUL = 'fullHangul'  # Full Width Hangul
    HALF_HANGUL = 'halfHangul'  # Half-Width Hangul IME Mode
    _members = (NO_CONTROL, OFF, IME_ON, DISABLED, HIRAGANA, FULL_KATAKANA, HALF_KATAKANA, FULL_ALPHA, HALF_ALPHA, FULL_HANGUL, HALF_HANGUL,)
    
class ST_CfType(XsdStringEnumeration):
    """
    Simple type: Conditional Format Type
    """
    EXPRESSION = 'expression'  # Expression
    CELL_IS = 'cellIs'  # Cell Is
    COLOR_SCALE = 'colorScale'  # Color Scale
    DATA_BAR = 'dataBar'  # Data Bar
    ICON_SET = 'iconSet'  # Icon Set
    TOP10 = 'top10'  # Top 10
    UNIQUE_VALUES = 'uniqueValues'  # Unique Values
    DUPLICATE_VALUES = 'duplicateValues'  # Duplicate Values
    CONTAINS_TEXT = 'containsText'  # Contains Text
    NOT_CONTAINS_TEXT = 'notContainsText'  # Does Not Contain Text
    BEGINS_WITH = 'beginsWith'  # Begins With
    ENDS_WITH = 'endsWith'  # Ends With
    CONTAINS_BLANKS = 'containsBlanks'  # Contains Blanks
    NOT_CONTAINS_BLANKS = 'notContainsBlanks'  # Contains No Blanks
    CONTAINS_ERRORS = 'containsErrors'  # Contains Errors
    NOT_CONTAINS_ERRORS = 'notContainsErrors'  # Contains No Errors
    TIME_PERIOD = 'timePeriod'  # Time Period
    ABOVE_AVERAGE = 'aboveAverage'  # Above or Below Average
    _members = (EXPRESSION, CELL_IS, COLOR_SCALE, DATA_BAR, ICON_SET, TOP10, UNIQUE_VALUES, DUPLICATE_VALUES, CONTAINS_TEXT, NOT_CONTAINS_TEXT, BEGINS_WITH, ENDS_WITH, CONTAINS_BLANKS, NOT_CONTAINS_BLANKS, CONTAINS_ERRORS, NOT_CONTAINS_ERRORS, TIME_PERIOD, ABOVE_AVERAGE,)
    
class ST_TimePeriod(XsdStringEnumeration):
    """
    Simple type: Time Period Types
    """
    TODAY = 'today'  # Today
    YESTERDAY = 'yesterday'  # Yesterday
    TOMORROW = 'tomorrow'  # Tomorrow
    LAST7_DAYS = 'last7Days'  # Last 7 Days
    THIS_MONTH = 'thisMonth'  # This Month
    LAST_MONTH = 'lastMonth'  # Last Month
    NEXT_MONTH = 'nextMonth'  # Next Month
    THIS_WEEK = 'thisWeek'  # This Week
    LAST_WEEK = 'lastWeek'  # Last Week
    NEXT_WEEK = 'nextWeek'  # Next Week
    _members = (TODAY, YESTERDAY, TOMORROW, LAST7_DAYS, THIS_MONTH, LAST_MONTH, NEXT_MONTH, THIS_WEEK, LAST_WEEK, NEXT_WEEK,)
    
class ST_ConditionalFormattingOperator(XsdStringEnumeration):
    """
    Simple type: Conditional Format Operators
    """
    LESS_THAN = 'lessThan'  # Less Than
    LESS_THAN_OR_EQUAL = 'lessThanOrEqual'  # Less Than Or Equal
    EQUAL = 'equal'  # Equal
    NOT_EQUAL = 'notEqual'  # Not Equal
    GREATER_THAN_OR_EQUAL = 'greaterThanOrEqual'  # Greater Than Or Equal
    GREATER_THAN = 'greaterThan'  # Greater Than
    BETWEEN = 'between'  # Between
    NOT_BETWEEN = 'notBetween'  # Not Between
    CONTAINS_TEXT = 'containsText'  # Contains
    NOT_CONTAINS = 'notContains'  # Does Not Contain
    BEGINS_WITH = 'beginsWith'  # Begins With
    ENDS_WITH = 'endsWith'  # Ends With
    _members = (LESS_THAN, LESS_THAN_OR_EQUAL, EQUAL, NOT_EQUAL, GREATER_THAN_OR_EQUAL, GREATER_THAN, BETWEEN, NOT_BETWEEN, CONTAINS_TEXT, NOT_CONTAINS, BEGINS_WITH, ENDS_WITH,)
    
class ST_CfvoType(XsdStringEnumeration):
    """
    Simple type: Conditional Format Value Object Type
    """
    NUM = 'num'  # Number
    PERCENT = 'percent'  # Percent
    MAX = 'max'  # Maximum
    MIN = 'min'  # Minimum
    FORMULA = 'formula'  # Formula
    PERCENTILE = 'percentile'  # Percentile
    _members = (NUM, PERCENT, MAX, MIN, FORMULA, PERCENTILE,)
    
class ST_PageOrder(XsdStringEnumeration):
    """
    Simple type: Page Order
    """
    DOWN_THEN_OVER = 'downThenOver'  # Down Then Over
    OVER_THEN_DOWN = 'overThenDown'  # Over Then Down
    _members = (DOWN_THEN_OVER, OVER_THEN_DOWN,)
    
class ST_Orientation(XsdStringEnumeration):
    """
    Simple type: Orientation
    """
    DEFAULT = 'default'  # Default
    PORTRAIT = 'portrait'  # Portrait
    LANDSCAPE = 'landscape'  # Landscape
    _members = (DEFAULT, PORTRAIT, LANDSCAPE,)
    
class ST_CellComments(XsdStringEnumeration):
    """
    Simple type: Cell Comments
    """
    NONE = 'none'  # None
    AS_DISPLAYED = 'asDisplayed'  # Print Comments As Displayed
    AT_END = 'atEnd'  # Print At End
    _members = (NONE, AS_DISPLAYED, AT_END,)
    
class ST_PrintError(XsdStringEnumeration):
    """
    Simple type: Print Errors
    """
    DISPLAYED = 'displayed'  # Display Cell Errors
    BLANK = 'blank'  # Show Cell Errors As Blank
    DASH = 'dash'  # Dash Cell Errors
    NA = 'NA'  # NA
    _members = (DISPLAYED, BLANK, DASH, NA,)
    
class ST_DvAspect(XsdStringEnumeration):
    """
    Simple type: Data View Aspect Type
    """
    DVASPECT_CONTENT = 'DVASPECT_CONTENT'  # Object Display Content
    DVASPECT_ICON = 'DVASPECT_ICON'  # Object Display Icon
    _members = (DVASPECT_CONTENT, DVASPECT_ICON,)
    
class ST_OleUpdate(XsdStringEnumeration):
    """
    Simple type: OLE Update Types
    """
    OLEUPDATE_ALWAYS = 'OLEUPDATE_ALWAYS'  # Always Update OLE
    OLEUPDATE_ONCALL = 'OLEUPDATE_ONCALL'  # Update OLE On Call
    _members = (OLEUPDATE_ALWAYS, OLEUPDATE_ONCALL,)
    
class ST_WebSourceType(XsdStringEnumeration):
    """
    Simple type: Web Source Type
    """
    SHEET = 'sheet'  # All Sheet Content
    PRINT_AREA = 'printArea'  # Print Area
    AUTO_FILTER = 'autoFilter'  # AutoFilter
    RANGE = 'range'  # Range
    CHART = 'chart'  # Chart
    PIVOT_TABLE = 'pivotTable'  # PivotTable
    QUERY = 'query'  # QueryTable
    LABEL = 'label'  # Label
    _members = (SHEET, PRINT_AREA, AUTO_FILTER, RANGE, CHART, PIVOT_TABLE, QUERY, LABEL,)
    
class ST_PaneState(XsdStringEnumeration):
    """
    Simple type: Pane State
    """
    SPLIT = 'split'  # Split
    FROZEN = 'frozen'  # Frozen
    FROZEN_SPLIT = 'frozenSplit'  # Frozen Split
    _members = (SPLIT, FROZEN, FROZEN_SPLIT,)
    
class ST_Xstring(XsdString):
    """
    Simple type: Escaped String
    """
    
class ST_CellRef(XsdString):
    """
    Simple type: Cell Reference
    """
    
class ST_Ref(XsdString):
    """
    Simple type: Cell References
    """
    
class ST_RefA(XsdString):
    """
    Simple type: Single Cell Reference
    """
    
class ST_Sqref():
    """
    Simple type: Reference Sequence
    """
    # list[itemType='ST_Ref']: None
    
class ST_Formula(ST_Xstring):
    """
    Simple type: Formula
    """
    
class ST_UnsignedIntHex(XsdHexBinary):
    """
    Simple type: Hex Unsigned Integer
    """
    # length[value='4']: None
    
class ST_UnsignedShortHex(XsdHexBinary):
    """
    Simple type: Unsigned Short Hex
    """
    # length[value='2']: None
    
class ST_Guid(XsdToken):
    """
    Simple type: Globally Unique Identifier
    """
    # pattern[value='\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}']: None
    
class ST_Visibility(XsdStringEnumeration):
    """
    Simple type: Visibility Types
    """
    VISIBLE = 'visible'  # Visible
    HIDDEN = 'hidden'  # Hidden
    VERY_HIDDEN = 'veryHidden'  # Very Hidden
    _members = (VISIBLE, HIDDEN, VERY_HIDDEN,)
    
class ST_Comments(XsdStringEnumeration):
    """
    Simple type: Comment Display Types
    """
    COMM_NONE = 'commNone'  # No Comments
    COMM_INDICATOR = 'commIndicator'  # Show Comment Indicator
    COMM_IND_AND_COMMENT = 'commIndAndComment'  # Show Comment & Indicator
    _members = (COMM_NONE, COMM_INDICATOR, COMM_IND_AND_COMMENT,)
    
class ST_Objects(XsdStringEnumeration):
    """
    Simple type: Object Display Types
    """
    ALL = 'all'  # All
    PLACEHOLDERS = 'placeholders'  # Show Placeholders
    NONE = 'none'  # None
    _members = (ALL, PLACEHOLDERS, NONE,)
    
class ST_SheetState(XsdStringEnumeration):
    """
    Simple type: Sheet Visibility Types
    """
    VISIBLE = 'visible'  # Visible
    HIDDEN = 'hidden'  # Hidden
    VERY_HIDDEN = 'veryHidden'  # Very Hidden
    _members = (VISIBLE, HIDDEN, VERY_HIDDEN,)
    
class ST_UpdateLinks(XsdStringEnumeration):
    """
    Simple type: Update Links Behavior Types
    """
    USER_SET = 'userSet'  # User Set
    NEVER = 'never'  # Never Update Links
    ALWAYS = 'always'  # Always Update Links
    _members = (USER_SET, NEVER, ALWAYS,)
    
class ST_SmartTagShow(XsdStringEnumeration):
    """
    Simple type: Smart Tag Display Types
    """
    ALL = 'all'  # All
    NONE = 'none'  # None
    NO_INDICATOR = 'noIndicator'  # No Smart Tag Indicator
    _members = (ALL, NONE, NO_INDICATOR,)
    
class ST_CalcMode(XsdStringEnumeration):
    """
    Simple type: Calculation Mode
    """
    MANUAL = 'manual'  # Manual Calculation Mode
    AUTO = 'auto'  # Automatic
    AUTO_NO_TABLE = 'autoNoTable'  # Automatic Calculation (No Tables)
    _members = (MANUAL, AUTO, AUTO_NO_TABLE,)
    
class ST_RefMode(XsdStringEnumeration):
    """
    Simple type: Reference Mode
    """
    A1_MODE = 'A1'  # A1 Mode
    R1_C1 = 'R1C1'  # R1C1 Reference Mode
    _members = (A1_MODE, R1_C1,)
    
class ST_TargetScreenSize(XsdStringEnumeration):
    """
    Simple type: Target Screen Size Types
    """
    # 544x376 544 x 376 Resolution
    # 640x480 640 x 480 Resolution
    # 720x512 720 x 512 Resolution
    # 800x600 800 x 600 Resolution
    # 1024x768 1024 x 768 Resolution
    # 1152x882 1152 x 882 Resolution
    # 1152x900 1152 x 900 Resolution
    # 1280x1024 1280 x 1024 Resolution
    # 1600x1200 1600 x 1200 Resolution
    # 1800x1440 1800 x 1440 Resolution
    # 1920x1200 1920 x 1200 Resolution
    _members = ('544x376', '640x480', '720x512', '800x600', '1024x768', '1152x882', '1152x900', '1280x1024', '1600x1200', '1800x1440', '1920x1200',)
    
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
    
class ST_FilterOperator(XsdStringEnumeration):
    """
    Simple type: Filter Operator
    """
    EQUAL = 'equal'  # Equal
    LESS_THAN = 'lessThan'  # Less Than
    LESS_THAN_OR_EQUAL = 'lessThanOrEqual'  # Less Than Or Equal
    NOT_EQUAL = 'notEqual'  # Not Equal
    GREATER_THAN_OR_EQUAL = 'greaterThanOrEqual'  # Greater Than Or Equal
    GREATER_THAN = 'greaterThan'  # Greater Than
    _members = (EQUAL, LESS_THAN, LESS_THAN_OR_EQUAL, NOT_EQUAL, GREATER_THAN_OR_EQUAL, GREATER_THAN,)
    
class ST_DynamicFilterType(XsdStringEnumeration):
    """
    Simple type: Dynamic Filter
    """
    NULL = 'null'  # Null
    ABOVE_AVERAGE = 'aboveAverage'  # Above Average
    BELOW_AVERAGE = 'belowAverage'  # Below Average
    TOMORROW = 'tomorrow'  # Tomorrow
    TODAY = 'today'  # Today
    YESTERDAY = 'yesterday'  # Yesterday
    NEXT_WEEK = 'nextWeek'  # Next Week
    THIS_WEEK = 'thisWeek'  # This Week
    LAST_WEEK = 'lastWeek'  # Last Week
    NEXT_MONTH = 'nextMonth'  # Next Month
    THIS_MONTH = 'thisMonth'  # This Month
    LAST_MONTH = 'lastMonth'  # Last Month
    NEXT_QUARTER = 'nextQuarter'  # Next Quarter
    THIS_QUARTER = 'thisQuarter'  # This Quarter
    LAST_QUARTER = 'lastQuarter'  # Last Quarter
    NEXT_YEAR = 'nextYear'  # Next Year
    THIS_YEAR = 'thisYear'  # This Year
    LAST_YEAR = 'lastYear'  # Last Year
    YEAR_TO_DATE = 'yearToDate'  # Year To Date
    # Q1 1st Quarter
    # Q2 2nd Quarter
    # Q3 3rd Quarter
    # Q4 4th Quarter
    # M1 1st Month
    # M2 2nd Month
    # M3 3rd Month
    # M4 4th Month
    # M5 5th Month
    # M6 6th Month
    # M7 7th Month
    # M8 8th Month
    # M9 9th Month
    M10 = 'M10'  # 10th Month
    M11 = 'M11'  # 11th Month
    M12 = 'M12'  # 12th Month
    _members = (NULL, ABOVE_AVERAGE, BELOW_AVERAGE, TOMORROW, TODAY, YESTERDAY, NEXT_WEEK, THIS_WEEK, LAST_WEEK, NEXT_MONTH, THIS_MONTH, LAST_MONTH, NEXT_QUARTER, THIS_QUARTER, LAST_QUARTER, NEXT_YEAR, THIS_YEAR, LAST_YEAR, YEAR_TO_DATE, 'Q1', 'Q2', 'Q3', 'Q4', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', M10, M11, M12,)
    
class ST_IconSetType(XsdStringEnumeration):
    """
    Simple type: Icon Set Type
    """
    # 3Arrows 3 Arrows
    # 3ArrowsGray 3 Arrows  (Gray)
    # 3Flags 3 Flags
    # 3TrafficLights1 3 Traffic Lights
    # 3TrafficLights2 3 Traffic Lights Black
    # 3Signs 3 Signs
    # 3Symbols 3 Symbols Circled
    # 3Symbols2 3 Symbols
    # 4Arrows 4 Arrows
    # 4ArrowsGray 4 Arrows (Gray)
    # 4RedToBlack 4 Red To Black
    # 4Rating 4 Ratings
    # 4TrafficLights 4 Traffic Lights
    # 5Arrows 5 Arrows
    # 5ArrowsGray 5 Arrows (Gray)
    # 5Rating 5 Ratings Icon Set
    # 5Quarters 5 Quarters
    _members = ('3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters',)
    
class ST_SortBy(XsdStringEnumeration):
    """
    Simple type: Sort By
    """
    VALUE = 'value'  # Value
    CELL_COLOR = 'cellColor'  # Sort by Cell Color
    FONT_COLOR = 'fontColor'  # Sort by Font Color
    ICON = 'icon'  # Sort by Icon
    _members = (VALUE, CELL_COLOR, FONT_COLOR, ICON,)
    
class ST_SortMethod(XsdStringEnumeration):
    """
    Simple type: Sort Method
    """
    STROKE = 'stroke'  # Sort by Stroke
    PIN_YIN = 'pinYin'  # PinYin Sort
    NONE = 'none'  # None
    _members = (STROKE, PIN_YIN, NONE,)
    
class ST_CalendarType(XsdStringEnumeration):
    """
    Simple type: Calendar Type
    """
    NONE = 'none'  # No Calendar Type
    GREGORIAN = 'gregorian'  # Gregorian
    GREGORIAN_US = 'gregorianUs'  # Gregorian (U.S.) Calendar
    JAPAN = 'japan'  # Japanese Emperor Era Calendar
    TAIWAN = 'taiwan'  # Taiwan Era Calendar
    KOREA = 'korea'  # Korean Tangun Era Calendar
    HIJRI = 'hijri'  # Hijri (Arabic Lunar) Calendar
    THAI = 'thai'  # Thai Calendar
    HEBREW = 'hebrew'  # Hebrew (Lunar) Calendar
    GREGORIAN_ME_FRENCH = 'gregorianMeFrench'  # Gregorian Middle East French Calendar
    GREGORIAN_ARABIC = 'gregorianArabic'  # Gregorian Arabic Calendar
    GREGORIAN_XLIT_ENGLISH = 'gregorianXlitEnglish'  # Gregorian Transliterated English Calendar
    GREGORIAN_XLIT_FRENCH = 'gregorianXlitFrench'  # Gregorian Transliterated French Calendar
    _members = (NONE, GREGORIAN, GREGORIAN_US, JAPAN, TAIWAN, KOREA, HIJRI, THAI, HEBREW, GREGORIAN_ME_FRENCH, GREGORIAN_ARABIC, GREGORIAN_XLIT_ENGLISH, GREGORIAN_XLIT_FRENCH,)
    
class ST_DateTimeGrouping(XsdStringEnumeration):
    """
    Simple type: Date Time Grouping
    """
    YEAR = 'year'  # Group by Year
    MONTH = 'month'  # Month
    DAY = 'day'  # Day
    HOUR = 'hour'  # Group by Hour
    MINUTE = 'minute'  # Group by Minute
    SECOND = 'second'  # Second
    _members = (YEAR, MONTH, DAY, HOUR, MINUTE, SECOND,)
    
class ST_PivotAreaType(XsdStringEnumeration):
    """
    Simple type: Rule Type
    """
    NONE = 'none'  # None
    NORMAL = 'normal'  # Normal
    DATA = 'data'  # Data
    ALL = 'all'  # All
    ORIGIN = 'origin'  # Origin
    BUTTON = 'button'  # Field Button
    TOP_RIGHT = 'topRight'  # Top Right
    _members = (NONE, NORMAL, DATA, ALL, ORIGIN, BUTTON, TOP_RIGHT,)
    
class ST_Axis(XsdStringEnumeration):
    """
    Simple type: PivotTable Axis
    """
    AXIS_ROW = 'axisRow'  # Row Axis
    AXIS_COL = 'axisCol'  # Column Axis
    AXIS_PAGE = 'axisPage'  # Include Count Filter
    AXIS_VALUES = 'axisValues'  # Values Axis
    _members = (AXIS_ROW, AXIS_COL, AXIS_PAGE, AXIS_VALUES,)

    



