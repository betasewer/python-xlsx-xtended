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
class CT_Worksheet(BaseOxmlElement):
    """
    Complex type: 
    """
    sheetPr = ZeroOrOne('ssml:sheetPr') # Worksheet Properties
    dimension = ZeroOrOne('ssml:dimension') # Worksheet Dimensions
    sheetViews = ZeroOrOne('ssml:sheetViews') # Sheet Views
    sheetFormatPr = ZeroOrOne('ssml:sheetFormatPr') # Sheet Format Properties
    cols = ZeroOrMore('ssml:cols') # Column Information
    sheetdata = OneAndOnlyOne('ssml:sheetData') # Sheet Data
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
    Complex type: 
    """
    row = ZeroOrMore('ssml:row') # Row
    
    
class CT_SheetCalcPr(BaseOxmlElement):
    """
    Complex type: 
    """
    fullCalcOnLoad = OptionalAttribute('ssml:fullCalcOnLoad', XsdBoolean) # Full Calculation On Load
    
class CT_SheetFormatPr(BaseOxmlElement):
    """
    Complex type: 
    """
    baseColWidth = OptionalAttribute('ssml:baseColWidth', XsdUnsignedInt) # Base Column Width
    defaultColWidth = OptionalAttribute('ssml:defaultColWidth', XsdDouble) # Default Column Width
    defaultRowHeight = RequiredAttribute('ssml:defaultRowHeight', XsdDouble) # Default Row Height
    customHeight = OptionalAttribute('ssml:customHeight', XsdBoolean) # Custom Height
    zeroHeight = OptionalAttribute('ssml:zeroHeight', XsdBoolean) # Hidden By Default
    thickTop = OptionalAttribute('ssml:thickTop', XsdBoolean) # Thick Top Border
    thickBottom = OptionalAttribute('ssml:thickBottom', XsdBoolean) # Thick Bottom Border
    outlineLevelRow = OptionalAttribute('ssml:outlineLevelRow', XsdUnsignedByte) # Maximum Outline Row
    outlineLevelCol = OptionalAttribute('ssml:outlineLevelCol', XsdUnsignedByte) # Column Outline Level
    
class CT_Cols(BaseOxmlElement):
    """
    Complex type: 
    """
    col = OneOrMore('ssml:col') # Column Width & Formatting
    
    
class CT_Col(BaseOxmlElement):
    """
    Complex type: 
    """
    min = RequiredAttribute('ssml:min', XsdUnsignedInt) # Minimum Column
    max = RequiredAttribute('ssml:max', XsdUnsignedInt) # Maximum Column
    width = OptionalAttribute('ssml:width', XsdDouble) # Column Width
    style = OptionalAttribute('ssml:style', XsdUnsignedInt) # Style
    hidden = OptionalAttribute('ssml:hidden', XsdBoolean) # Hidden Columns
    bestFit = OptionalAttribute('ssml:bestFit', XsdBoolean) # Best Fit Column Width
    customWidth = OptionalAttribute('ssml:customWidth', XsdBoolean) # Custom Width
    phonetic = OptionalAttribute('ssml:phonetic', XsdBoolean) # Show Phonetic Information
    outlineLevel = OptionalAttribute('ssml:outlineLevel', XsdUnsignedByte) # Outline Level
    collapsed = OptionalAttribute('ssml:collapsed', XsdBoolean) # Collapsed
    
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
    rich = ZeroOrOne('ssml:is') # Rich Text Inline
    #extLst = *ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    ref = OptionalAttribute('r', ST_CellRef) # Reference	
    style = OptionalAttribute('s', XsdUnsignedInt) # Style Index Default value is "0".
    type = OptionalAttribute('t', ST_CellType) #	Cell Data Type	Default value is "n".
    cell_metadata = OptionalAttribute('cm', XsdUnsignedInt) #	Cell Metadata Index	Default value is "0".
    value_metadata = OptionalAttribute('vm', XsdUnsignedInt) #	Value Metadata Index	Default value is "0".
    phonetic = OptionalAttribute('ph', XsdBoolean) #	Show Phonetic	Default value is "false".    
    
class CT_SheetPr(BaseOxmlElement):
    """
    Complex type: 
    """
    tabColor = ZeroOrOne('ssml:tabColor') # Sheet Tab Color
    outlinePr = ZeroOrOne('ssml:outlinePr') # Outline Properties
    pageSetUpPr = ZeroOrOne('ssml:pageSetUpPr') # Page Setup Properties
    
    syncHorizontal = OptionalAttribute('ssml:syncHorizontal', XsdBoolean) # Synch Horizontal
    syncVertical = OptionalAttribute('ssml:syncVertical', XsdBoolean) # Synch Vertical
    syncRef = OptionalAttribute('ssml:syncRef', ST_Ref) # Synch Reference
    transitionEvaluation = OptionalAttribute('ssml:transitionEvaluation', XsdBoolean) # Transition Formula Evaluation
    transitionEntry = OptionalAttribute('ssml:transitionEntry', XsdBoolean) # Transition Formula Entry
    published = OptionalAttribute('ssml:published', XsdBoolean) # Published
    codeName = OptionalAttribute('ssml:codeName', XsdString) # Code Name
    filterMode = OptionalAttribute('ssml:filterMode', XsdBoolean) # Filter Mode
    enableFormatConditionsCalculation = OptionalAttribute('ssml:enableFormatConditionsCalculation', XsdBoolean) # Enable Conditional Formatting Calculations
    
class CT_SheetDimension(BaseOxmlElement):
    """
    Complex type: 
    """
    ref = RequiredAttribute('ref', ST_Ref) # Reference
    
class CT_SheetViews(BaseOxmlElement):
    """
    Complex type: 
    """
    sheetView = OneOrMore('ssml:sheetView') # Worksheet View
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    
class CT_SheetView(BaseOxmlElement):
    """
    Complex type: 
    """
    pane = ZeroOrOne('ssml:pane') # View Pane
    selection = ZeroOrMore('ssml:selection') # [0..4] Selection
    pivotSelection = ZeroOrMore('ssml:pivotSelection') # [0..4] PivotTable Selection
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    windowProtection = OptionalAttribute('ssml:windowProtection', XsdBoolean) # Window Protection
    showFormulas = OptionalAttribute('ssml:showFormulas', XsdBoolean) # Show Formulas
    showGridLines = OptionalAttribute('ssml:showGridLines', XsdBoolean) # Show Grid Lines
    showRowColHeaders = OptionalAttribute('ssml:showRowColHeaders', XsdBoolean) # Show Headers
    showZeros = OptionalAttribute('ssml:showZeros', XsdBoolean) # Show Zero Values
    rightToLeft = OptionalAttribute('ssml:rightToLeft', XsdBoolean) # Right To Left
    tabSelected = OptionalAttribute('ssml:tabSelected', XsdBoolean) # Sheet Tab Selected
    showRuler = OptionalAttribute('ssml:showRuler', XsdBoolean) # Show Ruler
    showOutlineSymbols = OptionalAttribute('ssml:showOutlineSymbols', XsdBoolean) # Show Outline Symbols
    defaultGridColor = OptionalAttribute('ssml:defaultGridColor', XsdBoolean) # Default Grid Color
    showWhiteSpace = OptionalAttribute('ssml:showWhiteSpace', XsdBoolean) # Show White Space
    view = OptionalAttribute('ssml:view', ST_SheetViewType) # View Type
    topLeftCell = OptionalAttribute('ssml:topLeftCell', ST_CellRef) # Top Left Visible Cell
    colorId = OptionalAttribute('ssml:colorId', XsdUnsignedInt) # Color Id
    zoomScale = OptionalAttribute('ssml:zoomScale', XsdUnsignedInt) # Zoom Scale
    zoomScaleNormal = OptionalAttribute('ssml:zoomScaleNormal', XsdUnsignedInt) # Zoom Scale Normal View
    zoomScaleSheetLayoutView = OptionalAttribute('ssml:zoomScaleSheetLayoutView', XsdUnsignedInt) # Zoom Scale Page Break Preview
    zoomScalePageLayoutView = OptionalAttribute('ssml:zoomScalePageLayoutView', XsdUnsignedInt) # Zoom Scale Page Layout View
    workbookViewId = RequiredAttribute('ssml:workbookViewId', XsdUnsignedInt) # Workbook View Index
    
class CT_Pane(BaseOxmlElement):
    """
    Complex type: 
    """
    xSplit = OptionalAttribute('ssml:xSplit', XsdDouble) # Horizontal Split Position
    ySplit = OptionalAttribute('ssml:ySplit', XsdDouble) # Vertical Split Position
    topLeftCell = OptionalAttribute('ssml:topLeftCell', ST_CellRef) # Top Left Visible Cell
    activePane = OptionalAttribute('ssml:activePane', ST_Pane) # Active Pane
    state = OptionalAttribute('ssml:state', ST_PaneState) # Split State
    
class CT_PivotSelection(BaseOxmlElement):
    """
    Complex type: 
    """
    pivotArea = OneAndOnlyOne('ssml:pivotArea') # Pivot Area
    
    pane = OptionalAttribute('ssml:pane', ST_Pane) # Pane
    showHeader = OptionalAttribute('ssml:showHeader', XsdBoolean) # Show Header
    label = OptionalAttribute('ssml:label', XsdBoolean) # Label
    data = OptionalAttribute('ssml:data', XsdBoolean) # Data Selection
    extendable = OptionalAttribute('ssml:extendable', XsdBoolean) # Extendable
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Selection Count
    axis = OptionalAttribute('ssml:axis', ST_Axis) # Axis
    dimension = OptionalAttribute('ssml:dimension', XsdUnsignedInt) # Dimension
    start = OptionalAttribute('ssml:start', XsdUnsignedInt) # Start
    min = OptionalAttribute('ssml:min', XsdUnsignedInt) # Minimum
    max = OptionalAttribute('ssml:max', XsdUnsignedInt) # Maximum
    activeRow = OptionalAttribute('ssml:activeRow', XsdUnsignedInt) # Active Row
    activeCol = OptionalAttribute('ssml:activeCol', XsdUnsignedInt) # Active Column
    previousRow = OptionalAttribute('ssml:previousRow', XsdUnsignedInt) # Previous Row
    previousCol = OptionalAttribute('ssml:previousCol', XsdUnsignedInt) # Previous Column Selection
    click = OptionalAttribute('ssml:click', XsdUnsignedInt) # Click Count
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_Selection(BaseOxmlElement):
    """
    Complex type: 
    """
    pane = OptionalAttribute('ssml:pane', ST_Pane) # Pane
    activeCell = OptionalAttribute('ssml:activeCell', ST_CellRef) # Active Cell Location
    activeCellId = OptionalAttribute('ssml:activeCellId', XsdUnsignedInt) # Active Cell Index
    sqref = OptionalAttribute('ssml:sqref', ST_Sqref) # Sequence of References
    
class CT_PageBreak(BaseOxmlElement):
    """
    Complex type: 
    """
    brk = ZeroOrMore('ssml:brk') # Break
    
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Page Break Count
    manualBreakCount = OptionalAttribute('ssml:manualBreakCount', XsdUnsignedInt) # Manual Break Count
    
class CT_Break(BaseOxmlElement):
    """
    Complex type: 
    """
    id = OptionalAttribute('ssml:id', XsdUnsignedInt) # Id
    min = OptionalAttribute('ssml:min', XsdUnsignedInt) # Minimum
    max = OptionalAttribute('ssml:max', XsdUnsignedInt) # Maximum
    man = OptionalAttribute('ssml:man', XsdBoolean) # Manual Page Break
    pt = OptionalAttribute('ssml:pt', XsdBoolean) # Pivot-Created Page Break
    
class CT_OutlinePr(BaseOxmlElement):
    """
    Complex type: 
    """
    applyStyles = OptionalAttribute('ssml:applyStyles', XsdBoolean) # Apply Styles in Outline
    summaryBelow = OptionalAttribute('ssml:summaryBelow', XsdBoolean) # Summary Below
    summaryRight = OptionalAttribute('ssml:summaryRight', XsdBoolean) # Summary Right
    showOutlineSymbols = OptionalAttribute('ssml:showOutlineSymbols', XsdBoolean) # Show Outline Symbols
    
class CT_PageSetUpPr(BaseOxmlElement):
    """
    Complex type: 
    """
    autoPageBreaks = OptionalAttribute('ssml:autoPageBreaks', XsdBoolean) # Show Auto Page Breaks
    fitToPage = OptionalAttribute('ssml:fitToPage', XsdBoolean) # Fit To Page
    
class CT_DataConsolidate(BaseOxmlElement):
    """
    Complex type: 
    """
    dataRefs = ZeroOrOne('ssml:dataRefs') # Data Consolidation References
    
    function = OptionalAttribute('ssml:function', ST_DataConsolidateFunction) # Function Index
    leftLabels = OptionalAttribute('ssml:leftLabels', XsdBoolean) # Use Left Column Labels
    topLabels = OptionalAttribute('ssml:topLabels', XsdBoolean) # Labels In Top Row
    link = OptionalAttribute('ssml:link', XsdBoolean) # Link
    
class CT_DataRefs(BaseOxmlElement):
    """
    Complex type: 
    """
    dataRef = ZeroOrMore('ssml:dataRef') # Data Consolidation Reference
    
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Data Consolidation Reference Count
    
class CT_DataRef(BaseOxmlElement):
    """
    Complex type: 
    """
    ref = OptionalAttribute('ssml:ref', ST_Ref) # Reference
    name = OptionalAttribute('ssml:name', ST_Xstring) # Named Range
    sheet = OptionalAttribute('ssml:sheet', ST_Xstring) # Sheet Name
    rid = OptionalAttribute('r:id', ST_RelationshipId) # relationship Id
    
class CT_MergeCells(BaseOxmlElement):
    """
    Complex type: 
    """
    mergeCell = OneOrMore('ssml:mergeCell') # Merged Cell
    
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Count
    
class CT_MergeCell(BaseOxmlElement):
    """
    Complex type: 
    """
    ref = RequiredAttribute('ssml:ref', ST_Ref) # Reference
    
class CT_SmartTags(BaseOxmlElement):
    """
    Complex type: 
    """
    cellSmartTags = OneOrMore('ssml:cellSmartTags') # Cell Smart Tags
    
    
class CT_CellSmartTags(BaseOxmlElement):
    """
    Complex type: 
    """
    cellSmartTag = OneOrMore('ssml:cellSmartTag') # Cell Smart Tag
    
    r = RequiredAttribute('ssml:r', ST_CellRef) # Reference
    
class CT_CellSmartTag(BaseOxmlElement):
    """
    Complex type: 
    """
    cellSmartTagPr = ZeroOrMore('ssml:cellSmartTagPr') # Smart Tag Properties
    
    type = RequiredAttribute('ssml:type', XsdUnsignedInt) # Smart Tag Type Index
    deleted = OptionalAttribute('ssml:deleted', XsdBoolean) # Deleted
    xmlBased = OptionalAttribute('ssml:xmlBased', XsdBoolean) # XML Based
    
class CT_CellSmartTagPr(BaseOxmlElement):
    """
    Complex type: 
    """
    key = RequiredAttribute('ssml:key', ST_Xstring) # Key Name
    val = RequiredAttribute('ssml:val', ST_Xstring) # Value
    
class CT_Drawing(BaseOxmlElement):
    """
    Complex type: 
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship id
    
class CT_LegacyDrawing(BaseOxmlElement):
    """
    Complex type: 
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_CustomSheetViews(BaseOxmlElement):
    """
    Complex type: 
    """
    customSheetView = OneOrMore('ssml:customSheetView') # Custom Sheet View
    
    
class CT_CustomSheetView(BaseOxmlElement):
    """
    Complex type: 
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
    extLst = ZeroOrOne('ssml:extLst') # None
    
    guid = RequiredAttribute('ssml:guid', ST_Guid) # GUID
    scale = OptionalAttribute('ssml:scale', XsdUnsignedInt) # Print Scale
    colorId = OptionalAttribute('ssml:colorId', XsdUnsignedInt) # Color Id
    showPageBreaks = OptionalAttribute('ssml:showPageBreaks', XsdBoolean) # Show Page Breaks
    showFormulas = OptionalAttribute('ssml:showFormulas', XsdBoolean) # Show Formulas
    showGridLines = OptionalAttribute('ssml:showGridLines', XsdBoolean) # Show Grid Lines
    showRowCol = OptionalAttribute('ssml:showRowCol', XsdBoolean) # Show Headers
    outlineSymbols = OptionalAttribute('ssml:outlineSymbols', XsdBoolean) # Show Outline Symbols
    zeroValues = OptionalAttribute('ssml:zeroValues', XsdBoolean) # Show Zero Values
    fitToPage = OptionalAttribute('ssml:fitToPage', XsdBoolean) # Fit To Page
    printArea = OptionalAttribute('ssml:printArea', XsdBoolean) # Print Area Defined
    filter = OptionalAttribute('ssml:filter', XsdBoolean) # Filtered List
    showAutoFilter = OptionalAttribute('ssml:showAutoFilter', XsdBoolean) # Show AutoFitler Drop Down Controls
    hiddenRows = OptionalAttribute('ssml:hiddenRows', XsdBoolean) # Hidden Rows
    hiddenColumns = OptionalAttribute('ssml:hiddenColumns', XsdBoolean) # Hidden Columns
    state = OptionalAttribute('ssml:state', ST_SheetState) # Visible State
    filterUnique = OptionalAttribute('ssml:filterUnique', XsdBoolean) # Filter
    view = OptionalAttribute('ssml:view', ST_SheetViewType) # View Type
    showRuler = OptionalAttribute('ssml:showRuler', XsdBoolean) # Show Ruler
    topLeftCell = OptionalAttribute('ssml:topLeftCell', ST_CellRef) # Top Left Visible Cell
    
class CT_DataValidations(BaseOxmlElement):
    """
    Complex type: 
    """
    dataValidation = OneOrMore('ssml:dataValidation') # Data Validation
    
    disablePrompts = OptionalAttribute('ssml:disablePrompts', XsdBoolean) # Disable Prompts
    xWindow = OptionalAttribute('ssml:xWindow', XsdUnsignedInt) # Top Left Corner (X Coodrinate)
    yWindow = OptionalAttribute('ssml:yWindow', XsdUnsignedInt) # Top Left Corner (Y Coordinate)
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Data Validation Item Count
    
class CT_DataValidation(BaseOxmlElement):
    """
    Complex type: 
    """
    formula1 = ZeroOrOne('ssml:formula1') # Formula 1
    formula2 = ZeroOrOne('ssml:formula2') # Formula 2
    
    type = OptionalAttribute('ssml:type', ST_DataValidationType) # Data Validation Type
    errorStyle = OptionalAttribute('ssml:errorStyle', ST_DataValidationErrorStyle) # Data Validation Error Style
    imeMode = OptionalAttribute('ssml:imeMode', ST_DataValidationImeMode) # IME Mode Enforced
    operator = OptionalAttribute('ssml:operator', ST_DataValidationOperator) # Operator
    allowBlank = OptionalAttribute('ssml:allowBlank', XsdBoolean) # Allow Blank
    showDropDown = OptionalAttribute('ssml:showDropDown', XsdBoolean) # Show Drop Down
    showInputMessage = OptionalAttribute('ssml:showInputMessage', XsdBoolean) # Show Input Message
    showErrorMessage = OptionalAttribute('ssml:showErrorMessage', XsdBoolean) # Show Error Message
    errorTitle = OptionalAttribute('ssml:errorTitle', ST_Xstring) # Error Alert Text
    error = OptionalAttribute('ssml:error', ST_Xstring) # Error Message
    promptTitle = OptionalAttribute('ssml:promptTitle', ST_Xstring) # Prompt Title
    prompt = OptionalAttribute('ssml:prompt', ST_Xstring) # Input Prompt
    sqref = RequiredAttribute('ssml:sqref', ST_Sqref) # Sequence of References
    
class CT_ConditionalFormatting(BaseOxmlElement):
    """
    Complex type: 
    """
    cfRule = OneOrMore('ssml:cfRule') # Conditional Formatting Rule
    extLst = ZeroOrOne('ssml:extLst') # None
    
    pivot = OptionalAttribute('ssml:pivot', XsdBoolean) # PivotTable Conditional Formatting
    sqref = OptionalAttribute('ssml:sqref', ST_Sqref) # Sequence of Refernces
    
class CT_CfRule(BaseOxmlElement):
    """
    Complex type: 
    """
    formula = ZeroOrMore('ssml:formula') # Formula
    colorScale = ZeroOrOne('ssml:colorScale') # Color Scale
    dataBar = ZeroOrOne('ssml:dataBar') # Data Bar
    iconSet = ZeroOrOne('ssml:iconSet') # Icon Set
    extLst = ZeroOrOne('ssml:extLst') # None
    
    type = OptionalAttribute('ssml:type', ST_CfType) # Type
    dxfId = OptionalAttribute('ssml:dxfId', ST_DxfId) # Differential Formatting Id
    priority = RequiredAttribute('ssml:priority', XsdInt) # Priority
    stopIfTrue = OptionalAttribute('ssml:stopIfTrue', XsdBoolean) # Stop If True
    aboveAverage = OptionalAttribute('ssml:aboveAverage', XsdBoolean) # Above Or Below Average
    percent = OptionalAttribute('ssml:percent', XsdBoolean) # Top 10 Percent
    bottom = OptionalAttribute('ssml:bottom', XsdBoolean) # Bottom N
    operator = OptionalAttribute('ssml:operator', ST_ConditionalFormattingOperator) # Operator
    text = OptionalAttribute('ssml:text', XsdString) # Text
    timePeriod = OptionalAttribute('ssml:timePeriod', ST_TimePeriod) # Time Period
    rank = OptionalAttribute('ssml:rank', XsdUnsignedInt) # Rank
    stdDev = OptionalAttribute('ssml:stdDev', XsdInt) # StdDev
    equalAverage = OptionalAttribute('ssml:equalAverage', XsdBoolean) # Equal Average
    
class CT_Hyperlinks(BaseOxmlElement):
    """
    Complex type: 
    """
    hyperlink = OneOrMore('ssml:hyperlink') # Hyperlink
    
    
class CT_Hyperlink(BaseOxmlElement):
    """
    Complex type: 
    """
    ref = RequiredAttribute('ssml:ref', ST_Ref) # Reference
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Relationship Id
    location = OptionalAttribute('ssml:location', ST_Xstring) # Location
    tooltip = OptionalAttribute('ssml:tooltip', ST_Xstring) # Tool Tip
    display = OptionalAttribute('ssml:display', ST_Xstring) # Display String
    
class CT_CellFormula(BaseOxmlElement):
    """
    Complex type: 
    """
    # simpleContent[]: 
      
    
class CT_ColorScale(BaseOxmlElement):
    """
    Complex type: 
    """
    cfvo = OneOrMore('ssml:cfvo') # Conditional Format Value Object
    color = OneOrMore('ssml:color') # Color Gradiant Interpolation
    
    
class CT_DataBar(BaseOxmlElement):
    """
    Complex type: 
    """
    cfvo = OneOrMore('ssml:cfvo') # Conditional Format Value Object
    color = OneAndOnlyOne('ssml:color') # Data Bar Color
    
    minLength = OptionalAttribute('ssml:minLength', XsdUnsignedInt) # Minimum Length
    maxLength = OptionalAttribute('ssml:maxLength', XsdUnsignedInt) # Maximum Length
    showValue = OptionalAttribute('ssml:showValue', XsdBoolean) # Show Values
    
class CT_IconSet(BaseOxmlElement):
    """
    Complex type: 
    """
    cfvo = OneOrMore('ssml:cfvo') # Conditional Formatting Object
    
    iconSet = OptionalAttribute('ssml:iconSet', ST_IconSetType) # Icon Set
    showValue = OptionalAttribute('ssml:showValue', XsdBoolean) # Show Value
    percent = OptionalAttribute('ssml:percent', XsdBoolean) # Percent
    reverse = OptionalAttribute('ssml:reverse', XsdBoolean) # Reverse Icons
    
class CT_Cfvo(BaseOxmlElement):
    """
    Complex type: 
    """
    extLst = ZeroOrOne('ssml:extLst') # None
    
    type = RequiredAttribute('ssml:type', ST_CfvoType) # Type
    val = OptionalAttribute('ssml:val', ST_Xstring) # Value
    gte = OptionalAttribute('ssml:gte', XsdBoolean) # Greater Than Or Equal
    
class CT_PageMargins(BaseOxmlElement):
    """
    Complex type: 
    """
    left = RequiredAttribute('ssml:left', XsdDouble) # Left Page Margin
    right = RequiredAttribute('ssml:right', XsdDouble) # Right Page Margin
    top = RequiredAttribute('ssml:top', XsdDouble) # Top Page Margin
    bottom = RequiredAttribute('ssml:bottom', XsdDouble) # Bottom Page Margin
    header = RequiredAttribute('ssml:header', XsdDouble) # Header Page Margin
    footer = RequiredAttribute('ssml:footer', XsdDouble) # Footer Page Margin
    
class CT_PrintOptions(BaseOxmlElement):
    """
    Complex type: 
    """
    horizontalCentered = OptionalAttribute('ssml:horizontalCentered', XsdBoolean) # Horizontal Centered
    verticalCentered = OptionalAttribute('ssml:verticalCentered', XsdBoolean) # Vertical Centered
    headings = OptionalAttribute('ssml:headings', XsdBoolean) # Print Headings
    gridLines = OptionalAttribute('ssml:gridLines', XsdBoolean) # Print Grid Lines
    gridLinesSet = OptionalAttribute('ssml:gridLinesSet', XsdBoolean) # Grid Lines Set
    
class CT_PageSetup(BaseOxmlElement):
    """
    Complex type: 
    """
    paperSize = OptionalAttribute('ssml:paperSize', XsdUnsignedInt) # Paper Size
    scale = OptionalAttribute('ssml:scale', XsdUnsignedInt) # Print Scale
    firstPageNumber = OptionalAttribute('ssml:firstPageNumber', XsdUnsignedInt) # First Page Number
    fitToWidth = OptionalAttribute('ssml:fitToWidth', XsdUnsignedInt) # Fit To Width
    fitToHeight = OptionalAttribute('ssml:fitToHeight', XsdUnsignedInt) # Fit To Height
    pageOrder = OptionalAttribute('ssml:pageOrder', ST_PageOrder) # Page Order
    orientation = OptionalAttribute('ssml:orientation', ST_Orientation) # Orientation
    usePrinterDefaults = OptionalAttribute('ssml:usePrinterDefaults', XsdBoolean) # Use Printer Defaults
    blackAndWhite = OptionalAttribute('ssml:blackAndWhite', XsdBoolean) # Black And White
    draft = OptionalAttribute('ssml:draft', XsdBoolean) # Draft
    cellComments = OptionalAttribute('ssml:cellComments', ST_CellComments) # Print Cell Comments
    useFirstPageNumber = OptionalAttribute('ssml:useFirstPageNumber', XsdBoolean) # Use First Page Number
    errors = OptionalAttribute('ssml:errors', ST_PrintError) # Print Error Handling
    horizontalDpi = OptionalAttribute('ssml:horizontalDpi', XsdUnsignedInt) # Horizontal DPI
    verticalDpi = OptionalAttribute('ssml:verticalDpi', XsdUnsignedInt) # Vertical DPI
    copies = OptionalAttribute('ssml:copies', XsdUnsignedInt) # Number Of Copies
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Id
    
class CT_HeaderFooter(BaseOxmlElement):
    """
    Complex type: 
    """
    oddHeader = ZeroOrOne('ssml:oddHeader') # Odd Header
    oddFooter = ZeroOrOne('ssml:oddFooter') # Odd Page Footer
    evenHeader = ZeroOrOne('ssml:evenHeader') # Even Page Header
    evenFooter = ZeroOrOne('ssml:evenFooter') # Even Page Footer
    firstHeader = ZeroOrOne('ssml:firstHeader') # First Page Header
    firstFooter = ZeroOrOne('ssml:firstFooter') # First Page Footer
    
    differentOddEven = OptionalAttribute('ssml:differentOddEven', XsdBoolean) # Different Odd Even Header Footer
    differentFirst = OptionalAttribute('ssml:differentFirst', XsdBoolean) # Different First Page
    scaleWithDoc = OptionalAttribute('ssml:scaleWithDoc', XsdBoolean) # Scale Header & Footer With Document
    alignWithMargins = OptionalAttribute('ssml:alignWithMargins', XsdBoolean) # Align Margins
    
class CT_Scenarios(BaseOxmlElement):
    """
    Complex type: 
    """
    scenario = OneOrMore('ssml:scenario') # Scenario
    
    current = OptionalAttribute('ssml:current', XsdUnsignedInt) # Current Scenario
    show = OptionalAttribute('ssml:show', XsdUnsignedInt) # Last Shown Scenario
    sqref = OptionalAttribute('ssml:sqref', ST_Sqref) # Sequence of References
    
class CT_SheetProtection(BaseOxmlElement):
    """
    Complex type: 
    """
    password = OptionalAttribute('ssml:password', ST_UnsignedShortHex) # Password
    sheet = OptionalAttribute('ssml:sheet', XsdBoolean) # Sheet Locked
    objects = OptionalAttribute('ssml:objects', XsdBoolean) # Objects Locked
    scenarios = OptionalAttribute('ssml:scenarios', XsdBoolean) # Scenarios Locked
    formatCells = OptionalAttribute('ssml:formatCells', XsdBoolean) # Format Cells Locked
    formatColumns = OptionalAttribute('ssml:formatColumns', XsdBoolean) # Format Columns Locked
    formatRows = OptionalAttribute('ssml:formatRows', XsdBoolean) # Format Rows Locked
    insertColumns = OptionalAttribute('ssml:insertColumns', XsdBoolean) # Insert Columns Locked
    insertRows = OptionalAttribute('ssml:insertRows', XsdBoolean) # Insert Rows Locked
    insertHyperlinks = OptionalAttribute('ssml:insertHyperlinks', XsdBoolean) # Insert Hyperlinks Locked
    deleteColumns = OptionalAttribute('ssml:deleteColumns', XsdBoolean) # Delete Columns Locked
    deleteRows = OptionalAttribute('ssml:deleteRows', XsdBoolean) # Delete Rows Locked
    selectLockedCells = OptionalAttribute('ssml:selectLockedCells', XsdBoolean) # Select Locked Cells Locked
    sort = OptionalAttribute('ssml:sort', XsdBoolean) # Sort Locked
    autoFilter = OptionalAttribute('ssml:autoFilter', XsdBoolean) # AutoFilter Locked
    pivotTables = OptionalAttribute('ssml:pivotTables', XsdBoolean) # Pivot Tables Locked
    selectUnlockedCells = OptionalAttribute('ssml:selectUnlockedCells', XsdBoolean) # Select Unlocked Cells Locked
    
class CT_ProtectedRanges(BaseOxmlElement):
    """
    Complex type: 
    """
    protectedRange = OneOrMore('ssml:protectedRange') # Protected Range
    
    
class CT_ProtectedRange(BaseOxmlElement):
    """
    Complex type: 
    """
    password = OptionalAttribute('ssml:password', ST_UnsignedShortHex) # Password
    sqref = RequiredAttribute('ssml:sqref', ST_Sqref) # Sequence of References
    name = RequiredAttribute('ssml:name', ST_Xstring) # Name
    securityDescriptor = OptionalAttribute('ssml:securityDescriptor', XsdString) # Security Descriptor
    
class CT_Scenario(BaseOxmlElement):
    """
    Complex type: 
    """
    inputCells = OneOrMore('ssml:inputCells') # Input Cells
    
    name = RequiredAttribute('ssml:name', ST_Xstring) # Scenario Name
    locked = OptionalAttribute('ssml:locked', XsdBoolean) # Scenario Locked
    hidden = OptionalAttribute('ssml:hidden', XsdBoolean) # Hidden Scenario
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Changing Cell Count
    user = OptionalAttribute('ssml:user', ST_Xstring) # User Name
    comment = OptionalAttribute('ssml:comment', ST_Xstring) # Scenario Comment
    
class CT_InputCells(BaseOxmlElement):
    """
    Complex type: 
    """
    r = RequiredAttribute('ssml:r', ST_CellRef) # Reference
    deleted = OptionalAttribute('ssml:deleted', XsdBoolean) # Deleted
    undone = OptionalAttribute('ssml:undone', XsdBoolean) # Undone
    val = RequiredAttribute('ssml:val', ST_Xstring) # Value
    numFmtId = OptionalAttribute('ssml:numFmtId', ST_NumFmtId) # Number Format Id
    
class CT_CellWatches(BaseOxmlElement):
    """
    Complex type: 
    """
    cellWatch = OneOrMore('ssml:cellWatch') # Cell Watch Item
    
    
class CT_CellWatch(BaseOxmlElement):
    """
    Complex type: 
    """
    r = RequiredAttribute('ssml:r', ST_CellRef) # Reference
    
    
class CT_ChartsheetPr(BaseOxmlElement):
    """
    Complex type: 
    """
    tabColor = ZeroOrOne('ssml:tabColor') # None
    
    published = OptionalAttribute('ssml:published', XsdBoolean) # Published
    codeName = OptionalAttribute('ssml:codeName', XsdString) # Code Name
    
class CT_ChartsheetViews(BaseOxmlElement):
    """
    Complex type: 
    """
    sheetView = OneOrMore('ssml:sheetView') # Chart Sheet View
    extLst = ZeroOrOne('ssml:extLst') # None
    
    
class CT_ChartsheetView(BaseOxmlElement):
    """
    Complex type: 
    """
    extLst = ZeroOrOne('ssml:extLst') # None
    
    tabSelected = OptionalAttribute('ssml:tabSelected', XsdBoolean) # Sheet Tab Selected
    zoomScale = OptionalAttribute('ssml:zoomScale', XsdUnsignedInt) # Window Zoom Scale
    workbookViewId = RequiredAttribute('ssml:workbookViewId', XsdUnsignedInt) # Workbook View Id
    zoomToFit = OptionalAttribute('ssml:zoomToFit', XsdBoolean) # Zoom To Fit
    
class CT_ChartsheetProtection(BaseOxmlElement):
    """
    Complex type: 
    """
    password = OptionalAttribute('ssml:password', ST_UnsignedShortHex) # Password
    content = OptionalAttribute('ssml:content', XsdBoolean) # Contents
    objects = OptionalAttribute('ssml:objects', XsdBoolean) # Objects Locked
    
class CT_CsPageSetup(BaseOxmlElement):
    """
    Complex type: 
    """
    paperSize = OptionalAttribute('ssml:paperSize', XsdUnsignedInt) # Paper Size
    firstPageNumber = OptionalAttribute('ssml:firstPageNumber', XsdUnsignedInt) # First Page Number
    orientation = OptionalAttribute('ssml:orientation', ST_Orientation) # Orientation
    usePrinterDefaults = OptionalAttribute('ssml:usePrinterDefaults', XsdBoolean) # Use Printer Defaults
    blackAndWhite = OptionalAttribute('ssml:blackAndWhite', XsdBoolean) # Black And White
    draft = OptionalAttribute('ssml:draft', XsdBoolean) # Draft
    useFirstPageNumber = OptionalAttribute('ssml:useFirstPageNumber', XsdBoolean) # Use First Page Number
    horizontalDpi = OptionalAttribute('ssml:horizontalDpi', XsdUnsignedInt) # Horizontal DPI
    verticalDpi = OptionalAttribute('ssml:verticalDpi', XsdUnsignedInt) # Vertical DPI
    copies = OptionalAttribute('ssml:copies', XsdUnsignedInt) # Number Of Copies
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Id
    
class CT_CustomChartsheetViews(BaseOxmlElement):
    """
    Complex type: 
    """
    customSheetView = ZeroOrMore('ssml:customSheetView') # Custom Chart Sheet View
    
    
class CT_CustomChartsheetView(BaseOxmlElement):
    """
    Complex type: 
    """
    pageMargins = ZeroOrOne('ssml:pageMargins') # None
    pageSetup = ZeroOrOne('ssml:pageSetup') # Chart Sheet Page Setup
    headerFooter = ZeroOrOne('ssml:headerFooter') # None
    
    guid = RequiredAttribute('ssml:guid', ST_Guid) # GUID
    scale = OptionalAttribute('ssml:scale', XsdUnsignedInt) # Print Scale
    state = OptionalAttribute('ssml:state', ST_SheetState) # Visible State
    zoomToFit = OptionalAttribute('ssml:zoomToFit', XsdBoolean) # Zoom To Fit
    
class CT_CustomProperties(BaseOxmlElement):
    """
    Complex type: 
    """
    customPr = OneOrMore('ssml:customPr') # Custom Property
    
    
class CT_CustomProperty(BaseOxmlElement):
    """
    Complex type: 
    """
    name = RequiredAttribute('ssml:name', ST_Xstring) # Custom Property Name
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_OleObjects(BaseOxmlElement):
    """
    Complex type: 
    """
    oleObject = OneOrMore('ssml:oleObject') # OLE Object
    
    
class CT_OleObject(BaseOxmlElement):
    """
    Complex type: 
    """
    progId = OptionalAttribute('ssml:progId', XsdString) # OLE ProgId
    dvAspect = OptionalAttribute('ssml:dvAspect', ST_DvAspect) # Data or View Aspect
    link = OptionalAttribute('ssml:link', ST_Xstring) # OLE Link Moniker
    oleUpdate = OptionalAttribute('ssml:oleUpdate', ST_OleUpdate) # OLE Update
    autoLoad = OptionalAttribute('ssml:autoLoad', XsdBoolean) # Auto Load
    shapeId = RequiredAttribute('ssml:shapeId', XsdUnsignedInt) # Shape Id
    rid = OptionalAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_WebPublishItems(BaseOxmlElement):
    """
    Complex type: 
    """
    webPublishItem = OneOrMore('ssml:webPublishItem') # Web Publishing Item
    
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Web Publishing Items Count
    
class CT_WebPublishItem(BaseOxmlElement):
    """
    Complex type: 
    """
    id = RequiredAttribute('ssml:id', XsdUnsignedInt) # Id
    divId = RequiredAttribute('ssml:divId', ST_Xstring) # Destination Bookmark
    sourceType = RequiredAttribute('ssml:sourceType', ST_WebSourceType) # Web Source Type
    sourceRef = OptionalAttribute('ssml:sourceRef', ST_Ref) # Source Id
    sourceObject = OptionalAttribute('ssml:sourceObject', ST_Xstring) # Source Object Name
    destinationFile = RequiredAttribute('ssml:destinationFile', ST_Xstring) # Destination File Name
    title = OptionalAttribute('ssml:title', ST_Xstring) # Title
    autoRepublish = OptionalAttribute('ssml:autoRepublish', XsdBoolean) # Automatically Publish
    
class CT_Controls(BaseOxmlElement):
    """
    Complex type: 
    """
    control = OneOrMore('ssml:control') # Embedded Control
    
    
class CT_Control(BaseOxmlElement):
    """
    Complex type: 
    """
    shapeId = RequiredAttribute('ssml:shapeId', XsdUnsignedInt) # Shape Id
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    name = OptionalAttribute('ssml:name', XsdString) # Control Name
    
class CT_IgnoredErrors(BaseOxmlElement):
    """
    Complex type: 
    """
    ignoredError = OneOrMore('ssml:ignoredError') # Ignored Error
    extLst = ZeroOrOne('ssml:extLst') # None
    
    
class CT_IgnoredError(BaseOxmlElement):
    """
    Complex type: 
    """
    sqref = RequiredAttribute('ssml:sqref', ST_Sqref) # Sequence of References
    evalError = OptionalAttribute('ssml:evalError', XsdBoolean) # Evaluation Error
    twoDigitTextYear = OptionalAttribute('ssml:twoDigitTextYear', XsdBoolean) # Two Digit Text Year
    numberStoredAsText = OptionalAttribute('ssml:numberStoredAsText', XsdBoolean) # Number Stored As Text
    formula = OptionalAttribute('ssml:formula', XsdBoolean) # Formula
    formulaRange = OptionalAttribute('ssml:formulaRange', XsdBoolean) # Formula Range
    unlockedFormula = OptionalAttribute('ssml:unlockedFormula', XsdBoolean) # Unlocked Formula
    emptyCellReference = OptionalAttribute('ssml:emptyCellReference', XsdBoolean) # Empty Cell Reference
    listDataValidation = OptionalAttribute('ssml:listDataValidation', XsdBoolean) # List Data Validation
    calculatedColumn = OptionalAttribute('ssml:calculatedColumn', XsdBoolean) # Calculated Column
    
class CT_TableParts(BaseOxmlElement):
    """
    Complex type: 
    """
    tablePart = ZeroOrMore('ssml:tablePart') # Table Part
    
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Count
    
class CT_TablePart(BaseOxmlElement):
    """
    Complex type: 
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
