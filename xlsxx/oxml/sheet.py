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
    XsdUnsignedInt, ST_RelationshipId, XsdDouble, XsdBoolean, XsdUnsignedByte, XsdString, XsdInt
)
from xlsxx.oxml.simpletypes import (
    ST_CellComments, ST_CellRef, ST_CellSpans, ST_CellType, ST_CfType, ST_CfvoType, 
    ST_ConditionalFormattingOperator, ST_DataConsolidateFunction, ST_DataValidationErrorStyle, 
    ST_DataValidationImeMode, ST_DataValidationOperator, ST_DataValidationType, ST_DvAspect, 
    ST_DxfId, ST_Guid, ST_NumFmtId, ST_OleUpdate, ST_Orientation, ST_PageOrder, ST_Pane, 
    ST_PaneState, ST_PrintError, ST_Ref, ST_SheetState, ST_SheetViewType, ST_Sqref, ST_TimePeriod, 
    ST_UnsignedShortHex, ST_WebSourceType, ST_Xstring,
)
# add implement later
ST_Axis = ST_Xstring
ST_IconSetType = ST_Xstring

#
#
#
#
#
#
class CT_Macrosheet(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    sheetPr = ZeroOrOne('ssml:sheetPr') # Sheet Properties
    dimension = ZeroOrOne('ssml:dimension') # Macro Sheet Dimensions
    sheetViews = ZeroOrOne('ssml:sheetViews') # Macro Sheet Views
    sheetFormatPr = ZeroOrOne('ssml:sheetFormatPr') # Sheet Format Properties
    cols = ZeroOrMore('ssml:cols') # Column Information
    sheetData = OneAndOnlyOne('ssml:sheetData') # Sheet Data
    sheetProtection = ZeroOrOne('ssml:sheetProtection') # Sheet Protection Options
    autoFilter = ZeroOrOne('ssml:autoFilter') # AutoFilter
    sortState = ZeroOrOne('ssml:sortState') # Sort State
    dataConsolidate = ZeroOrOne('ssml:dataConsolidate') # Data Consolidation
    customSheetViews = ZeroOrOne('ssml:customSheetViews') # Custom Sheet Views
    phoneticPr = ZeroOrOne('ssml:phoneticPr') # Phonetic Properties
    conditionalFormatting = ZeroOrMore('ssml:conditionalFormatting') # Conditional Formatting
    printOptions = ZeroOrOne('ssml:printOptions') # Print Options
    pageMargins = ZeroOrOne('ssml:pageMargins') # Page Margins
    pageSetup = ZeroOrOne('ssml:pageSetup') # Page Setup Settings
    headerFooter = ZeroOrOne('ssml:headerFooter') # Header Footer Settings
    rowBreaks = ZeroOrOne('ssml:rowBreaks') # Horizontal Page Breaks (Row)
    colBreaks = ZeroOrOne('ssml:colBreaks') # Vertical Page Breaks
    customProperties = ZeroOrOne('ssml:customProperties') # Custom Properties
    drawing = ZeroOrOne('ssml:drawing') # Drawing
    legacyDrawing = ZeroOrOne('ssml:legacyDrawing') # Legacy Drawing Reference
    legacyDrawingHF = ZeroOrOne('ssml:legacyDrawingHF') # Legacy Drawing Header Footer
    picture = ZeroOrOne('ssml:picture') # Background Image
    oleObjects = ZeroOrOne('ssml:oleObjects') # OLE Objects
    extLst = ZeroOrMore('ssml:extLst') # Future Feature Data Storage Area
    
    
class CT_Dialogsheet(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    sheetPr = ZeroOrMore('ssml:sheetPr') # Sheet Properties
    sheetViews = ZeroOrMore('ssml:sheetViews') # Dialog Sheet Views
    sheetFormatPr = ZeroOrMore('ssml:sheetFormatPr') # Dialog Sheet Format Properties
    sheetProtection = ZeroOrOne('ssml:sheetProtection') # Sheet Protection
    customSheetViews = ZeroOrMore('ssml:customSheetViews') # Custom Sheet Views
    printOptions = ZeroOrMore('ssml:printOptions') # Print Options
    pageMargins = ZeroOrMore('ssml:pageMargins') # Page Margins
    pageSetup = ZeroOrMore('ssml:pageSetup') # Page Setup Settings
    headerFooter = ZeroOrMore('ssml:headerFooter') # Header & Footer Settings
    drawing = ZeroOrMore('ssml:drawing') # Drawing
    legacyDrawing = ZeroOrMore('ssml:legacyDrawing') # Legacy Drawing
    legacyDrawingHF = ZeroOrOne('ssml:legacyDrawingHF') # Legacy Drawing Header Footer
    oleObjects = ZeroOrOne('ssml:oleObjects') # None
    extLst = ZeroOrMore('ssml:extLst') # Future Feature Data Storage Area
    
    
class CT_Worksheet(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    sheetPr = ZeroOrOne('ssml:sheetPr') # Worksheet Properties
    dimension = ZeroOrOne('ssml:dimension') # Worksheet Dimensions
    sheetViews = ZeroOrOne('ssml:sheetViews') # Sheet Views
    sheetFormatPr = ZeroOrOne('ssml:sheetFormatPr') # Sheet Format Properties
    cols = ZeroOrMore('ssml:cols') # Column Information
    sheetData = OneAndOnlyOne('ssml:sheetData') # Sheet Data
    sheetCalcPr = ZeroOrOne('ssml:sheetCalcPr') # Sheet Calculation Properties
    sheetProtection = ZeroOrOne('ssml:sheetProtection') # Sheet Protection
    protectedRanges = ZeroOrOne('ssml:protectedRanges') # Protected Ranges
    scenarios = ZeroOrOne('ssml:scenarios') # Scenarios
    autoFilter = ZeroOrOne('ssml:autoFilter') # AutoFilter
    sortState = ZeroOrOne('ssml:sortState') # Sort State
    dataConsolidate = ZeroOrOne('ssml:dataConsolidate') # Data Consolidate
    customSheetViews = ZeroOrOne('ssml:customSheetViews') # Custom Sheet Views
    mergeCells = ZeroOrOne('ssml:mergeCells') # Merge Cells
    phoneticPr = ZeroOrOne('ssml:phoneticPr') # Phonetic Properties
    conditionalFormatting = ZeroOrMore('ssml:conditionalFormatting') # Conditional Formatting
    dataValidations = ZeroOrOne('ssml:dataValidations') # Data Validations
    hyperlinks = ZeroOrOne('ssml:hyperlinks') # Hyperlinks
    printOptions = ZeroOrOne('ssml:printOptions') # Print Options
    pageMargins = ZeroOrOne('ssml:pageMargins') # Page Margins
    pageSetup = ZeroOrOne('ssml:pageSetup') # Page Setup Settings
    headerFooter = ZeroOrOne('ssml:headerFooter') # Header and Footer Settings
    rowBreaks = ZeroOrOne('ssml:rowBreaks') # Horizontal Page Breaks
    colBreaks = ZeroOrOne('ssml:colBreaks') # Vertical Page Breaks
    customProperties = ZeroOrOne('ssml:customProperties') # Custom Properties
    cellWatches = ZeroOrOne('ssml:cellWatches') # Cell Watch Items
    ignoredErrors = ZeroOrOne('ssml:ignoredErrors') # Ignored Errors
    smartTags = ZeroOrOne('ssml:smartTags') # Smart Tags
    drawing = ZeroOrOne('ssml:drawing') # Drawing
    legacyDrawing = ZeroOrOne('ssml:legacyDrawing') # Legacy Drawing
    legacyDrawingHF = ZeroOrOne('ssml:legacyDrawingHF') # Legacy Drawing Header Footer
    picture = ZeroOrOne('ssml:picture') # Background Image
    oleObjects = ZeroOrOne('ssml:oleObjects') # None
    controls = ZeroOrOne('ssml:controls') # Embedded Controls
    webPublishItems = ZeroOrOne('ssml:webPublishItems') # Web Publishing Items
    tableParts = ZeroOrOne('ssml:tableParts') # Table Parts
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    
class CT_SheetData(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    row = ZeroOrMore('ssml:row') # Row
    
    
class CT_SheetCalcPr(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    fullCalcOnLoad = OptionalAttribute('fullCalcOnLoad', XsdBoolean, False) # Full Calculation On Load
    
class CT_SheetFormatPr(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    baseColWidth = OptionalAttribute('baseColWidth', XsdUnsignedInt, 8) # Base Column Width
    defaultColWidth = OptionalAttribute('defaultColWidth', XsdDouble) # Default Column Width
    defaultRowHeight = RequiredAttribute('defaultRowHeight', XsdDouble) # Default Row Height
    customHeight = OptionalAttribute('customHeight', XsdBoolean, False) # Custom Height
    zeroHeight = OptionalAttribute('zeroHeight', XsdBoolean, False) # Hidden By Default
    thickTop = OptionalAttribute('thickTop', XsdBoolean, False) # Thick Top Border
    thickBottom = OptionalAttribute('thickBottom', XsdBoolean, False) # Thick Bottom Border
    outlineLevelRow = OptionalAttribute('outlineLevelRow', XsdUnsignedByte, 0) # Maximum Outline Row
    outlineLevelCol = OptionalAttribute('outlineLevelCol', XsdUnsignedByte, 0) # Column Outline Level
    
class CT_Cols(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    col = OneOrMore('ssml:col') # Column Width & Formatting
    
    
class CT_Col(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    min = RequiredAttribute('min', XsdUnsignedInt) # Minimum Column
    max = RequiredAttribute('max', XsdUnsignedInt) # Maximum Column
    width = OptionalAttribute('width', XsdDouble) # Column Width
    style = OptionalAttribute('style', XsdUnsignedInt, 0) # Style
    hidden = OptionalAttribute('hidden', XsdBoolean, False) # Hidden Columns
    bestFit = OptionalAttribute('bestFit', XsdBoolean, False) # Best Fit Column Width
    customWidth = OptionalAttribute('customWidth', XsdBoolean, False) # Custom Width
    phonetic = OptionalAttribute('phonetic', XsdBoolean, False) # Show Phonetic Information
    outlineLevel = OptionalAttribute('outlineLevel', XsdUnsignedByte, 0) # Outline Level
    collapsed = OptionalAttribute('collapsed', XsdBoolean, False) # Collapsed
    
class CT_Row(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    c = ZeroOrMore('ssml:c') # Cell
    extLst = ZeroOrMore('ssml:extLst') # Future Feature Data Storage Area
    
    r = OptionalAttribute('r', XsdUnsignedInt) # Row Index
    spans = OptionalAttribute('spans', ST_CellSpans) # Spans
    s = OptionalAttribute('s', XsdUnsignedInt, 0) # Style Index
    customFormat = OptionalAttribute('customFormat', XsdBoolean, False) # Custom Format
    ht = OptionalAttribute('ht', XsdDouble) # Row Height
    hidden = OptionalAttribute('hidden', XsdBoolean, False) # Hidden
    customHeight = OptionalAttribute('customHeight', XsdBoolean, False) # Custom Height
    outlineLevel = OptionalAttribute('outlineLevel', XsdUnsignedByte, 0) # Outline Level
    collapsed = OptionalAttribute('collapsed', XsdBoolean, False) # Collapsed
    thickTop = OptionalAttribute('thickTop', XsdBoolean, False) # Thick Top Border
    thickBot = OptionalAttribute('thickBot', XsdBoolean, False) # Thick Bottom
    ph = OptionalAttribute('ph', XsdBoolean, False) # Show Phonetic
    
class CT_Cell(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    f = ZeroOrOne('ssml:f') # Formula
    v = ZeroOrOne('ssml:v') # Cell Value
    is_ = ZeroOrOne('ssml:is') # Rich Text Inline
    extLst = ZeroOrMore('ssml:extLst') # Future Feature Data Storage Area
    
    r = OptionalAttribute('r', ST_CellRef) # Reference
    s = OptionalAttribute('s', XsdUnsignedInt, 0) # Style Index
    t = OptionalAttribute('t', ST_CellType, 'n') # Cell Data Type
    cm = OptionalAttribute('cm', XsdUnsignedInt, 0) # Cell Metadata Index
    vm = OptionalAttribute('vm', XsdUnsignedInt, 0) # Value Metadata Index
    ph = OptionalAttribute('ph', XsdBoolean, False) # Show Phonetic
    
class CT_SheetPr(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    tabColor = ZeroOrOne('ssml:tabColor') # Sheet Tab Color
    outlinePr = ZeroOrOne('ssml:outlinePr') # Outline Properties
    pageSetUpPr = ZeroOrOne('ssml:pageSetUpPr') # Page Setup Properties
    
    syncHorizontal = OptionalAttribute('syncHorizontal', XsdBoolean, False) # Synch Horizontal
    syncVertical = OptionalAttribute('syncVertical', XsdBoolean, False) # Synch Vertical
    syncRef = OptionalAttribute('syncRef', ST_Ref) # Synch Reference
    transitionEvaluation = OptionalAttribute('transitionEvaluation', XsdBoolean, False) # Transition Formula Evaluation
    transitionEntry = OptionalAttribute('transitionEntry', XsdBoolean, False) # Transition Formula Entry
    published = OptionalAttribute('published', XsdBoolean, True) # Published
    codeName = OptionalAttribute('codeName', XsdString) # Code Name
    filterMode = OptionalAttribute('filterMode', XsdBoolean, False) # Filter Mode
    enableFormatConditionsCalculation = OptionalAttribute('enableFormatConditionsCalculation', XsdBoolean, True) # Enable Conditional Formatting Calculations
    
class CT_SheetDimension(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    ref = RequiredAttribute('ref', ST_Ref) # Reference
    
class CT_SheetViews(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    sheetView = OneOrMore('ssml:sheetView') # Worksheet View
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    
class CT_SheetView(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    pane = ZeroOrOne('ssml:pane') # View Pane
    selection = ZeroOrMore('ssml:selection') # Selection
    pivotSelection = ZeroOrMore('ssml:pivotSelection') # PivotTable Selection
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    windowProtection = OptionalAttribute('windowProtection', XsdBoolean, False) # Window Protection
    showFormulas = OptionalAttribute('showFormulas', XsdBoolean, False) # Show Formulas
    showGridLines = OptionalAttribute('showGridLines', XsdBoolean, True) # Show Grid Lines
    showRowColHeaders = OptionalAttribute('showRowColHeaders', XsdBoolean, True) # Show Headers
    showZeros = OptionalAttribute('showZeros', XsdBoolean, True) # Show Zero Values
    rightToLeft = OptionalAttribute('rightToLeft', XsdBoolean, False) # Right To Left
    tabSelected = OptionalAttribute('tabSelected', XsdBoolean, False) # Sheet Tab Selected
    showRuler = OptionalAttribute('showRuler', XsdBoolean, True) # Show Ruler
    showOutlineSymbols = OptionalAttribute('showOutlineSymbols', XsdBoolean, True) # Show Outline Symbols
    defaultGridColor = OptionalAttribute('defaultGridColor', XsdBoolean, True) # Default Grid Color
    showWhiteSpace = OptionalAttribute('showWhiteSpace', XsdBoolean, True) # Show White Space
    view = OptionalAttribute('view', ST_SheetViewType, 'normal') # View Type
    topLeftCell = OptionalAttribute('topLeftCell', ST_CellRef) # Top Left Visible Cell
    colorId = OptionalAttribute('colorId', XsdUnsignedInt, 64) # Color Id
    zoomScale = OptionalAttribute('zoomScale', XsdUnsignedInt, 100) # Zoom Scale
    zoomScaleNormal = OptionalAttribute('zoomScaleNormal', XsdUnsignedInt, 0) # Zoom Scale Normal View
    zoomScaleSheetLayoutView = OptionalAttribute('zoomScaleSheetLayoutView', XsdUnsignedInt, 0) # Zoom Scale Page Break Preview
    zoomScalePageLayoutView = OptionalAttribute('zoomScalePageLayoutView', XsdUnsignedInt, 0) # Zoom Scale Page Layout View
    workbookViewId = RequiredAttribute('workbookViewId', XsdUnsignedInt) # Workbook View Index
    
class CT_Pane(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    xSplit = OptionalAttribute('xSplit', XsdDouble, 0.0) # Horizontal Split Position
    ySplit = OptionalAttribute('ySplit', XsdDouble, 0.0) # Vertical Split Position
    topLeftCell = OptionalAttribute('topLeftCell', ST_CellRef) # Top Left Visible Cell
    activePane = OptionalAttribute('activePane', ST_Pane, 'topLeft') # Active Pane
    state = OptionalAttribute('state', ST_PaneState, 'split') # Split State
    
class CT_PivotSelection(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    pivotArea = OneAndOnlyOne('ssml:pivotArea') # Pivot Area
    
    pane = OptionalAttribute('pane', ST_Pane, 'topLeft') # Pane
    showHeader = OptionalAttribute('showHeader', XsdBoolean) # Show Header
    label = OptionalAttribute('label', XsdBoolean) # Label
    data = OptionalAttribute('data', XsdBoolean) # Data Selection
    extendable = OptionalAttribute('extendable', XsdBoolean) # Extendable
    count = OptionalAttribute('count', XsdUnsignedInt) # Selection Count
    axis = OptionalAttribute('axis', ST_Axis) # Axis
    dimension = OptionalAttribute('dimension', XsdUnsignedInt) # Dimension
    start = OptionalAttribute('start', XsdUnsignedInt) # Start
    min = OptionalAttribute('min', XsdUnsignedInt) # Minimum
    max = OptionalAttribute('max', XsdUnsignedInt) # Maximum
    activeRow = OptionalAttribute('activeRow', XsdUnsignedInt) # Active Row
    activeCol = OptionalAttribute('activeCol', XsdUnsignedInt) # Active Column
    previousRow = OptionalAttribute('previousRow', XsdUnsignedInt) # Previous Row
    previousCol = OptionalAttribute('previousCol', XsdUnsignedInt) # Previous Column Selection
    click = OptionalAttribute('click', XsdUnsignedInt) # Click Count
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_Selection(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    pane = OptionalAttribute('pane', ST_Pane, 'topLeft') # Pane
    activeCell = OptionalAttribute('activeCell', ST_CellRef) # Active Cell Location
    activeCellId = OptionalAttribute('activeCellId', XsdUnsignedInt, 0) # Active Cell Index
    sqref = OptionalAttribute('sqref', ST_Sqref, 'A1') # Sequence of References
    
class CT_PageBreak(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    brk = ZeroOrMore('ssml:brk') # Break
    
    count = OptionalAttribute('count', XsdUnsignedInt, 0) # Page Break Count
    manualBreakCount = OptionalAttribute('manualBreakCount', XsdUnsignedInt, 0) # Manual Break Count
    
class CT_Break(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    id = OptionalAttribute('id', XsdUnsignedInt, 0) # Id
    min = OptionalAttribute('min', XsdUnsignedInt, 0) # Minimum
    max = OptionalAttribute('max', XsdUnsignedInt, 0) # Maximum
    man = OptionalAttribute('man', XsdBoolean, False) # Manual Page Break
    pt = OptionalAttribute('pt', XsdBoolean, False) # Pivot-Created Page Break
    
class CT_OutlinePr(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    applyStyles = OptionalAttribute('applyStyles', XsdBoolean, False) # Apply Styles in Outline
    summaryBelow = OptionalAttribute('summaryBelow', XsdBoolean, True) # Summary Below
    summaryRight = OptionalAttribute('summaryRight', XsdBoolean, True) # Summary Right
    showOutlineSymbols = OptionalAttribute('showOutlineSymbols', XsdBoolean, True) # Show Outline Symbols
    
class CT_PageSetUpPr(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    autoPageBreaks = OptionalAttribute('autoPageBreaks', XsdBoolean, True) # Show Auto Page Breaks
    fitToPage = OptionalAttribute('fitToPage', XsdBoolean, False) # Fit To Page
    
class CT_DataConsolidate(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    dataRefs = ZeroOrOne('ssml:dataRefs') # Data Consolidation References
    
    function = OptionalAttribute('function', ST_DataConsolidateFunction, 'sum') # Function Index
    leftLabels = OptionalAttribute('leftLabels', XsdBoolean, False) # Use Left Column Labels
    topLabels = OptionalAttribute('topLabels', XsdBoolean, False) # Labels In Top Row
    link = OptionalAttribute('link', XsdBoolean, False) # Link
    
class CT_DataRefs(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    dataRef = ZeroOrMore('ssml:dataRef') # Data Consolidation Reference
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Data Consolidation Reference Count
    
class CT_DataRef(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    ref = OptionalAttribute('ref', ST_Ref) # Reference
    name = OptionalAttribute('name', ST_Xstring) # Named Range
    sheet = OptionalAttribute('sheet', ST_Xstring) # Sheet Name
    rid = OptionalAttribute('r:id', ST_RelationshipId) # relationship Id
    
class CT_MergeCells(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    mergeCell = OneOrMore('ssml:mergeCell') # Merged Cell
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Count
    
class CT_MergeCell(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    ref = RequiredAttribute('ref', ST_Ref) # Reference
    
class CT_SmartTags(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    cellSmartTags = OneOrMore('ssml:cellSmartTags') # Cell Smart Tags
    
    
class CT_CellSmartTags(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    cellSmartTag = OneOrMore('ssml:cellSmartTag') # Cell Smart Tag
    
    r = RequiredAttribute('r', ST_CellRef) # Reference
    
class CT_CellSmartTag(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    cellSmartTagPr = ZeroOrMore('ssml:cellSmartTagPr') # Smart Tag Properties
    
    type = RequiredAttribute('type', XsdUnsignedInt) # Smart Tag Type Index
    deleted = OptionalAttribute('deleted', XsdBoolean, False) # Deleted
    xmlBased = OptionalAttribute('xmlBased', XsdBoolean, False) # XML Based
    
class CT_CellSmartTagPr(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    key = RequiredAttribute('key', ST_Xstring) # Key Name
    val = RequiredAttribute('val', ST_Xstring) # Value
    
class CT_Drawing(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship id
    
class CT_LegacyDrawing(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_CustomSheetViews(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    customSheetView = OneOrMore('ssml:customSheetView') # Custom Sheet View
    
    
class CT_CustomSheetView(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    pane = ZeroOrOne('ssml:pane') # Pane Split Information
    selection = ZeroOrOne('ssml:selection') # Selection
    rowBreaks = ZeroOrOne('ssml:rowBreaks') # Horizontal Page Breaks
    colBreaks = ZeroOrOne('ssml:colBreaks') # Vertical Page Breaks
    pageMargins = ZeroOrOne('ssml:pageMargins') # Page Margins
    printOptions = ZeroOrOne('ssml:printOptions') # Print Options
    pageSetup = ZeroOrOne('ssml:pageSetup') # Page Setup Settings
    headerFooter = ZeroOrOne('ssml:headerFooter') # Header Footer Settings
    autoFilter = ZeroOrOne('ssml:autoFilter') # AutoFilter Settings
    extLst = ZeroOrMore('ssml:extLst') # None
    
    guid = RequiredAttribute('guid', ST_Guid) # GUID
    scale = OptionalAttribute('scale', XsdUnsignedInt) # Print Scale
    colorId = OptionalAttribute('colorId', XsdUnsignedInt) # Color Id
    showPageBreaks = OptionalAttribute('showPageBreaks', XsdBoolean, False) # Show Page Breaks
    showFormulas = OptionalAttribute('showFormulas', XsdBoolean, False) # Show Formulas
    showGridLines = OptionalAttribute('showGridLines', XsdBoolean, True) # Show Grid Lines
    showRowCol = OptionalAttribute('showRowCol', XsdBoolean, True) # Show Headers
    outlineSymbols = OptionalAttribute('outlineSymbols', XsdBoolean, True) # Show Outline Symbols
    zeroValues = OptionalAttribute('zeroValues', XsdBoolean, True) # Show Zero Values
    fitToPage = OptionalAttribute('fitToPage', XsdBoolean, False) # Fit To Page
    printArea = OptionalAttribute('printArea', XsdBoolean, False) # Print Area Defined
    filter = OptionalAttribute('filter', XsdBoolean, False) # Filtered List
    showAutoFilter = OptionalAttribute('showAutoFilter', XsdBoolean, False) # Show AutoFitler Drop Down Controls
    hiddenRows = OptionalAttribute('hiddenRows', XsdBoolean, False) # Hidden Rows
    hiddenColumns = OptionalAttribute('hiddenColumns', XsdBoolean, False) # Hidden Columns
    state = OptionalAttribute('state', ST_SheetState) # Visible State
    filterUnique = OptionalAttribute('filterUnique', XsdBoolean, False) # Filter
    view = OptionalAttribute('view', ST_SheetViewType) # View Type
    showRuler = OptionalAttribute('showRuler', XsdBoolean, True) # Show Ruler
    topLeftCell = OptionalAttribute('topLeftCell', ST_CellRef) # Top Left Visible Cell
    
class CT_DataValidations(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    dataValidation = OneOrMore('ssml:dataValidation') # Data Validation
    
    disablePrompts = OptionalAttribute('disablePrompts', XsdBoolean, False) # Disable Prompts
    xWindow = OptionalAttribute('xWindow', XsdUnsignedInt) # Top Left Corner (X Coodrinate)
    yWindow = OptionalAttribute('yWindow', XsdUnsignedInt) # Top Left Corner (Y Coordinate)
    count = OptionalAttribute('count', XsdUnsignedInt) # Data Validation Item Count
    
class CT_DataValidation(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    formula1 = ZeroOrOne('ssml:formula1') # Formula 1
    formula2 = ZeroOrOne('ssml:formula2') # Formula 2
    
    type = OptionalAttribute('type', ST_DataValidationType, 'none') # Data Validation Type
    errorStyle = OptionalAttribute('errorStyle', ST_DataValidationErrorStyle, 'stop') # Data Validation Error Style
    imeMode = OptionalAttribute('imeMode', ST_DataValidationImeMode, 'noControl') # IME Mode Enforced
    operator = OptionalAttribute('operator', ST_DataValidationOperator, 'between') # Operator
    allowBlank = OptionalAttribute('allowBlank', XsdBoolean, False) # Allow Blank
    showDropDown = OptionalAttribute('showDropDown', XsdBoolean, False) # Show Drop Down
    showInputMessage = OptionalAttribute('showInputMessage', XsdBoolean, False) # Show Input Message
    showErrorMessage = OptionalAttribute('showErrorMessage', XsdBoolean, False) # Show Error Message
    errorTitle = OptionalAttribute('errorTitle', ST_Xstring) # Error Alert Text
    error = OptionalAttribute('error', ST_Xstring) # Error Message
    promptTitle = OptionalAttribute('promptTitle', ST_Xstring) # Prompt Title
    prompt = OptionalAttribute('prompt', ST_Xstring) # Input Prompt
    sqref = RequiredAttribute('sqref', ST_Sqref) # Sequence of References
    
class CT_ConditionalFormatting(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    cfRule = OneOrMore('ssml:cfRule') # Conditional Formatting Rule
    extLst = ZeroOrMore('ssml:extLst') # None
    
    pivot = OptionalAttribute('pivot', XsdBoolean) # PivotTable Conditional Formatting
    sqref = OptionalAttribute('sqref', ST_Sqref) # Sequence of Refernces
    
class CT_CfRule(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    formula = ZeroOrMore('ssml:formula') # Formula
    colorScale = ZeroOrOne('ssml:colorScale') # Color Scale
    dataBar = ZeroOrOne('ssml:dataBar') # Data Bar
    iconSet = ZeroOrOne('ssml:iconSet') # Icon Set
    extLst = ZeroOrMore('ssml:extLst') # None
    
    type = OptionalAttribute('type', ST_CfType) # Type
    dxfId = OptionalAttribute('dxfId', ST_DxfId) # Differential Formatting Id
    priority = RequiredAttribute('priority', XsdInt) # Priority
    stopIfTrue = OptionalAttribute('stopIfTrue', XsdBoolean, False) # Stop If True
    aboveAverage = OptionalAttribute('aboveAverage', XsdBoolean, True) # Above Or Below Average
    percent = OptionalAttribute('percent', XsdBoolean, False) # Top 10 Percent
    bottom = OptionalAttribute('bottom', XsdBoolean, False) # Bottom N
    operator = OptionalAttribute('operator', ST_ConditionalFormattingOperator) # Operator
    text = OptionalAttribute('text', XsdString) # Text
    timePeriod = OptionalAttribute('timePeriod', ST_TimePeriod) # Time Period
    rank = OptionalAttribute('rank', XsdUnsignedInt) # Rank
    stdDev = OptionalAttribute('stdDev', XsdInt) # StdDev
    equalAverage = OptionalAttribute('equalAverage', XsdBoolean, False) # Equal Average
    
class CT_Hyperlinks(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    hyperlink = OneOrMore('ssml:hyperlink') # Hyperlink
    
    
class CT_Hyperlink(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    ref = RequiredAttribute('ref', ST_Ref) # Reference
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Relationship Id
    location = OptionalAttribute('location', ST_Xstring) # Location
    tooltip = OptionalAttribute('tooltip', ST_Xstring) # Tool Tip
    display = OptionalAttribute('display', ST_Xstring) # Display String
    
class CT_CellFormula(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    # simpleContent[]: 
      
    
class CT_ColorScale(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    cfvo = ZeroOrMore('ssml:cfvo') # Conditional Format Value Object
    color = ZeroOrMore('ssml:color') # Color Gradiant Interpolation
    
    
class CT_DataBar(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    cfvo = ZeroOrMore('ssml:cfvo') # Conditional Format Value Object
    color = OneAndOnlyOne('ssml:color') # Data Bar Color
    
    minLength = OptionalAttribute('minLength', XsdUnsignedInt, 10) # Minimum Length
    maxLength = OptionalAttribute('maxLength', XsdUnsignedInt, 90) # Maximum Length
    showValue = OptionalAttribute('showValue', XsdBoolean, True) # Show Values
    
class CT_IconSet(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    cfvo = ZeroOrMore('ssml:cfvo') # Conditional Formatting Object
    
    iconSet = OptionalAttribute('iconSet', ST_IconSetType, '3TrafficLights1') # Icon Set
    showValue = OptionalAttribute('showValue', XsdBoolean, True) # Show Value
    percent = OptionalAttribute('percent', XsdBoolean) # Percent
    reverse = OptionalAttribute('reverse', XsdBoolean, False) # Reverse Icons
    
class CT_Cfvo(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    extLst = ZeroOrOne('ssml:extLst') # None
    
    type = RequiredAttribute('type', ST_CfvoType) # Type
    val = OptionalAttribute('val', ST_Xstring) # Value
    gte = OptionalAttribute('gte', XsdBoolean, True) # Greater Than Or Equal
    
class CT_PageMargins(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    left = RequiredAttribute('left', XsdDouble) # Left Page Margin
    right = RequiredAttribute('right', XsdDouble) # Right Page Margin
    top = RequiredAttribute('top', XsdDouble) # Top Page Margin
    bottom = RequiredAttribute('bottom', XsdDouble) # Bottom Page Margin
    header = RequiredAttribute('header', XsdDouble) # Header Page Margin
    footer = RequiredAttribute('footer', XsdDouble) # Footer Page Margin
    
class CT_PrintOptions(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    horizontalCentered = OptionalAttribute('horizontalCentered', XsdBoolean, False) # Horizontal Centered
    verticalCentered = OptionalAttribute('verticalCentered', XsdBoolean, False) # Vertical Centered
    headings = OptionalAttribute('headings', XsdBoolean, False) # Print Headings
    gridLines = OptionalAttribute('gridLines', XsdBoolean, False) # Print Grid Lines
    gridLinesSet = OptionalAttribute('gridLinesSet', XsdBoolean, True) # Grid Lines Set
    
class CT_PageSetup(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    paperSize = OptionalAttribute('paperSize', XsdUnsignedInt, 1) # Paper Size
    scale = OptionalAttribute('scale', XsdUnsignedInt, 100) # Print Scale
    firstPageNumber = OptionalAttribute('firstPageNumber', XsdUnsignedInt, 1) # First Page Number
    fitToWidth = OptionalAttribute('fitToWidth', XsdUnsignedInt, 1) # Fit To Width
    fitToHeight = OptionalAttribute('fitToHeight', XsdUnsignedInt, 1) # Fit To Height
    pageOrder = OptionalAttribute('pageOrder', ST_PageOrder, 'downThenOver') # Page Order
    orientation = OptionalAttribute('orientation', ST_Orientation, 'default') # Orientation
    usePrinterDefaults = OptionalAttribute('usePrinterDefaults', XsdBoolean, True) # Use Printer Defaults
    blackAndWhite = OptionalAttribute('blackAndWhite', XsdBoolean, False) # Black And White
    draft = OptionalAttribute('draft', XsdBoolean, False) # Draft
    cellComments = OptionalAttribute('cellComments', ST_CellComments, 'none') # Print Cell Comments
    useFirstPageNumber = OptionalAttribute('useFirstPageNumber', XsdBoolean, False) # Use First Page Number
    errors = OptionalAttribute('errors', ST_PrintError, 'displayed') # Print Error Handling
    horizontalDpi = OptionalAttribute('horizontalDpi', XsdUnsignedInt, 600) # Horizontal DPI
    verticalDpi = OptionalAttribute('verticalDpi', XsdUnsignedInt, 600) # Vertical DPI
    copies = OptionalAttribute('copies', XsdUnsignedInt, 1) # Number Of Copies
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Id
    
class CT_HeaderFooter(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    oddHeader = ZeroOrOne('ssml:oddHeader') # Odd Header
    oddFooter = ZeroOrOne('ssml:oddFooter') # Odd Page Footer
    evenHeader = ZeroOrOne('ssml:evenHeader') # Even Page Header
    evenFooter = ZeroOrOne('ssml:evenFooter') # Even Page Footer
    firstHeader = ZeroOrOne('ssml:firstHeader') # First Page Header
    firstFooter = ZeroOrOne('ssml:firstFooter') # First Page Footer
    
    differentOddEven = OptionalAttribute('differentOddEven', XsdBoolean) # Different Odd Even Header Footer
    differentFirst = OptionalAttribute('differentFirst', XsdBoolean) # Different First Page
    scaleWithDoc = OptionalAttribute('scaleWithDoc', XsdBoolean) # Scale Header & Footer With Document
    alignWithMargins = OptionalAttribute('alignWithMargins', XsdBoolean) # Align Margins
    
class CT_Scenarios(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    scenario = OneOrMore('ssml:scenario') # Scenario
    
    current = OptionalAttribute('current', XsdUnsignedInt) # Current Scenario
    show = OptionalAttribute('show', XsdUnsignedInt) # Last Shown Scenario
    sqref = OptionalAttribute('sqref', ST_Sqref) # Sequence of References
    
class CT_SheetProtection(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    password = OptionalAttribute('password', ST_UnsignedShortHex) # Password
    sheet = OptionalAttribute('sheet', XsdBoolean, False) # Sheet Locked
    objects = OptionalAttribute('objects', XsdBoolean, False) # Objects Locked
    scenarios = OptionalAttribute('scenarios', XsdBoolean, False) # Scenarios Locked
    formatCells = OptionalAttribute('formatCells', XsdBoolean, True) # Format Cells Locked
    formatColumns = OptionalAttribute('formatColumns', XsdBoolean, True) # Format Columns Locked
    formatRows = OptionalAttribute('formatRows', XsdBoolean, True) # Format Rows Locked
    insertColumns = OptionalAttribute('insertColumns', XsdBoolean, True) # Insert Columns Locked
    insertRows = OptionalAttribute('insertRows', XsdBoolean, True) # Insert Rows Locked
    insertHyperlinks = OptionalAttribute('insertHyperlinks', XsdBoolean, True) # Insert Hyperlinks Locked
    deleteColumns = OptionalAttribute('deleteColumns', XsdBoolean, True) # Delete Columns Locked
    deleteRows = OptionalAttribute('deleteRows', XsdBoolean, True) # Delete Rows Locked
    selectLockedCells = OptionalAttribute('selectLockedCells', XsdBoolean, False) # Select Locked Cells Locked
    sort = OptionalAttribute('sort', XsdBoolean, True) # Sort Locked
    autoFilter = OptionalAttribute('autoFilter', XsdBoolean, True) # AutoFilter Locked
    pivotTables = OptionalAttribute('pivotTables', XsdBoolean, True) # Pivot Tables Locked
    selectUnlockedCells = OptionalAttribute('selectUnlockedCells', XsdBoolean, False) # Select Unlocked Cells Locked
    
class CT_ProtectedRanges(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    protectedRange = OneOrMore('ssml:protectedRange') # Protected Range
    
    
class CT_ProtectedRange(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    password = OptionalAttribute('password', ST_UnsignedShortHex) # Password
    sqref = RequiredAttribute('sqref', ST_Sqref) # Sequence of References
    name = RequiredAttribute('name', ST_Xstring) # Name
    securityDescriptor = OptionalAttribute('securityDescriptor', XsdString) # Security Descriptor
    
class CT_Scenario(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    inputCells = OneOrMore('ssml:inputCells') # Input Cells
    
    name = RequiredAttribute('name', ST_Xstring) # Scenario Name
    locked = OptionalAttribute('locked', XsdBoolean, False) # Scenario Locked
    hidden = OptionalAttribute('hidden', XsdBoolean, False) # Hidden Scenario
    count = OptionalAttribute('count', XsdUnsignedInt) # Changing Cell Count
    user = OptionalAttribute('user', ST_Xstring) # User Name
    comment = OptionalAttribute('comment', ST_Xstring) # Scenario Comment
    
class CT_InputCells(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    r = RequiredAttribute('r', ST_CellRef) # Reference
    deleted = OptionalAttribute('deleted', XsdBoolean, False) # Deleted
    undone = OptionalAttribute('undone', XsdBoolean, False) # Undone
    val = RequiredAttribute('val', ST_Xstring) # Value
    numFmtId = OptionalAttribute('numFmtId', ST_NumFmtId) # Number Format Id
    
class CT_CellWatches(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    cellWatch = OneOrMore('ssml:cellWatch') # Cell Watch Item
    
    
class CT_CellWatch(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    r = RequiredAttribute('r', ST_CellRef) # Reference
    
class CT_Chartsheet(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    sheetPr = ZeroOrOne('ssml:sheetPr') # Chart Sheet Properties
    sheetViews = OneAndOnlyOne('ssml:sheetViews') # Chart Sheet Views
    sheetProtection = ZeroOrOne('ssml:sheetProtection') # Chart Sheet Protection
    customSheetViews = ZeroOrOne('ssml:customSheetViews') # Custom Chart Sheet Views
    pageMargins = ZeroOrMore('ssml:pageMargins') # None
    pageSetup = ZeroOrOne('ssml:pageSetup') # None
    headerFooter = ZeroOrMore('ssml:headerFooter') # None
    drawing = OneAndOnlyOne('ssml:drawing') # Drawing
    legacyDrawing = ZeroOrOne('ssml:legacyDrawing') # None
    legacyDrawingHF = ZeroOrOne('ssml:legacyDrawingHF') # Legacy Drawing Reference in  Header Footer
    picture = ZeroOrOne('ssml:picture') # None
    webPublishItems = ZeroOrOne('ssml:webPublishItems') # None
    extLst = ZeroOrOne('ssml:extLst') # None
    
    
class CT_ChartsheetPr(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    tabColor = ZeroOrOne('ssml:tabColor') # None
    
    published = OptionalAttribute('published', XsdBoolean, True) # Published
    codeName = OptionalAttribute('codeName', XsdString) # Code Name
    
class CT_ChartsheetViews(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    sheetView = OneOrMore('ssml:sheetView') # Chart Sheet View
    extLst = ZeroOrOne('ssml:extLst') # None
    
    
class CT_ChartsheetView(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    extLst = ZeroOrOne('ssml:extLst') # None
    
    tabSelected = OptionalAttribute('tabSelected', XsdBoolean, False) # Sheet Tab Selected
    zoomScale = OptionalAttribute('zoomScale', XsdUnsignedInt, 100) # Window Zoom Scale
    workbookViewId = RequiredAttribute('workbookViewId', XsdUnsignedInt) # Workbook View Id
    zoomToFit = OptionalAttribute('zoomToFit', XsdBoolean, False) # Zoom To Fit
    
class CT_ChartsheetProtection(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    password = OptionalAttribute('password', ST_UnsignedShortHex) # Password
    content = OptionalAttribute('content', XsdBoolean, False) # Contents
    objects = OptionalAttribute('objects', XsdBoolean, False) # Objects Locked
    
class CT_CsPageSetup(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    paperSize = OptionalAttribute('paperSize', XsdUnsignedInt, 1) # Paper Size
    firstPageNumber = OptionalAttribute('firstPageNumber', XsdUnsignedInt, 1) # First Page Number
    orientation = OptionalAttribute('orientation', ST_Orientation, 'default') # Orientation
    usePrinterDefaults = OptionalAttribute('usePrinterDefaults', XsdBoolean, True) # Use Printer Defaults
    blackAndWhite = OptionalAttribute('blackAndWhite', XsdBoolean, False) # Black And White
    draft = OptionalAttribute('draft', XsdBoolean, False) # Draft
    useFirstPageNumber = OptionalAttribute('useFirstPageNumber', XsdBoolean, False) # Use First Page Number
    horizontalDpi = OptionalAttribute('horizontalDpi', XsdUnsignedInt, 600) # Horizontal DPI
    verticalDpi = OptionalAttribute('verticalDpi', XsdUnsignedInt, 600) # Vertical DPI
    copies = OptionalAttribute('copies', XsdUnsignedInt, 1) # Number Of Copies
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Id
    
class CT_CustomChartsheetViews(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    customSheetView = ZeroOrMore('ssml:customSheetView') # Custom Chart Sheet View
    
    
class CT_CustomChartsheetView(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    pageMargins = ZeroOrOne('ssml:pageMargins') # None
    pageSetup = ZeroOrOne('ssml:pageSetup') # Chart Sheet Page Setup
    headerFooter = ZeroOrOne('ssml:headerFooter') # None
    
    guid = RequiredAttribute('guid', ST_Guid) # GUID
    scale = OptionalAttribute('scale', XsdUnsignedInt) # Print Scale
    state = OptionalAttribute('state', ST_SheetState) # Visible State
    zoomToFit = OptionalAttribute('zoomToFit', XsdBoolean, False) # Zoom To Fit
    
class CT_CustomProperties(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    customPr = OneOrMore('ssml:customPr') # Custom Property
    
    
class CT_CustomProperty(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    name = RequiredAttribute('name', ST_Xstring) # Custom Property Name
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_OleObjects(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    oleObject = OneOrMore('ssml:oleObject') # OLE Object
    
    
class CT_OleObject(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    progId = OptionalAttribute('progId', XsdString) # OLE ProgId
    dvAspect = OptionalAttribute('dvAspect', ST_DvAspect, 'DVASPECT_CONTENT') # Data or View Aspect
    link = OptionalAttribute('link', ST_Xstring) # OLE Link Moniker
    oleUpdate = OptionalAttribute('oleUpdate', ST_OleUpdate) # OLE Update
    autoLoad = OptionalAttribute('autoLoad', XsdBoolean, False) # Auto Load
    shapeId = RequiredAttribute('shapeId', XsdUnsignedInt) # Shape Id
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_WebPublishItems(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    webPublishItem = OneOrMore('ssml:webPublishItem') # Web Publishing Item
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Web Publishing Items Count
    
class CT_WebPublishItem(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    id = RequiredAttribute('id', XsdUnsignedInt) # Id
    divId = RequiredAttribute('divId', ST_Xstring) # Destination Bookmark
    sourceType = RequiredAttribute('sourceType', ST_WebSourceType) # Web Source Type
    sourceRef = OptionalAttribute('sourceRef', ST_Ref) # Source Id
    sourceObject = OptionalAttribute('sourceObject', ST_Xstring) # Source Object Name
    destinationFile = RequiredAttribute('destinationFile', ST_Xstring) # Destination File Name
    title = OptionalAttribute('title', ST_Xstring) # Title
    autoRepublish = OptionalAttribute('autoRepublish', XsdBoolean, False) # Automatically Publish
    
class CT_Controls(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    control = OneOrMore('ssml:control') # Embedded Control
    
    
class CT_Control(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    shapeId = RequiredAttribute('shapeId', XsdUnsignedInt) # Shape Id
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    name = OptionalAttribute('name', XsdString) # Control Name
    
class CT_IgnoredErrors(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    ignoredError = OneOrMore('ssml:ignoredError') # Ignored Error
    extLst = ZeroOrOne('ssml:extLst') # None
    
    
class CT_IgnoredError(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    sqref = RequiredAttribute('sqref', ST_Sqref) # Sequence of References
    evalError = OptionalAttribute('evalError', XsdBoolean, False) # Evaluation Error
    twoDigitTextYear = OptionalAttribute('twoDigitTextYear', XsdBoolean, False) # Two Digit Text Year
    numberStoredAsText = OptionalAttribute('numberStoredAsText', XsdBoolean, False) # Number Stored As Text
    formula = OptionalAttribute('formula', XsdBoolean, False) # Formula
    formulaRange = OptionalAttribute('formulaRange', XsdBoolean, False) # Formula Range
    unlockedFormula = OptionalAttribute('unlockedFormula', XsdBoolean, False) # Unlocked Formula
    emptyCellReference = OptionalAttribute('emptyCellReference', XsdBoolean, False) # Empty Cell Reference
    listDataValidation = OptionalAttribute('listDataValidation', XsdBoolean, False) # List Data Validation
    calculatedColumn = OptionalAttribute('calculatedColumn', XsdBoolean, False) # Calculated Column
    
class CT_TableParts(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    tablePart = ZeroOrMore('ssml:tablePart') # Table Part
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Count
    
class CT_TablePart(BaseOxmlElement):
    """
    Complex type (sml-sheet.xsd)
    
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
