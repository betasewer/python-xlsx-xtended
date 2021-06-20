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
    XsdUnsignedInt, ST_RelationshipId, XsdToken, XsdBoolean, XsdInt, XsdDouble, XsdString
)
from xlsxx.oxml.simpletypes import (
    ST_Xstring, ST_Guid, ST_Visibility, ST_Comments, ST_SheetState, ST_Objects, ST_UnsignedShortHex,
    ST_UpdateLinks, ST_SmartTagShow, ST_CalcMode, ST_RefMode, ST_Ref, ST_TargetScreenSize
)

class CT_Workbook(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    fileVersion = ZeroOrOne('ssml:fileVersion') # File Version
    fileSharing = ZeroOrOne('ssml:fileSharing') # File Sharing
    workbookPr = ZeroOrOne('ssml:workbookPr') # Workbook Properties
    workbookProtection = ZeroOrOne('ssml:workbookProtection') # Workbook Protection
    bookViews = ZeroOrOne('ssml:bookViews') # Workbook Views
    sheets = OneAndOnlyOne('ssml:sheets') # Sheets
    functionGroups = ZeroOrOne('ssml:functionGroups') # Function Groups
    externalReferences = ZeroOrOne('ssml:externalReferences') # External References
    definedNames = ZeroOrOne('ssml:definedNames') # Defined Names
    calcPr = ZeroOrOne('ssml:calcPr') # Calculation Properties
    oleSize = ZeroOrOne('ssml:oleSize') # OLE Size
    customWorkbookViews = ZeroOrOne('ssml:customWorkbookViews') # Custom Workbook Views
    pivotCaches = ZeroOrOne('ssml:pivotCaches') # PivotCaches
    smartTagPr = ZeroOrOne('ssml:smartTagPr') # Smart Tag Properties
    smartTagTypes = ZeroOrOne('ssml:smartTagTypes') # Smart Tag Types
    webPublishing = ZeroOrOne('ssml:webPublishing') # Web Publishing Properties
    fileRecoveryPr = ZeroOrMore('ssml:fileRecoveryPr') # File Recovery Properties
    webPublishObjects = ZeroOrOne('ssml:webPublishObjects') # Web Publish Objects
    extLst = ZeroOrOne('ssml:extLst') # Future Feature Data Storage Area
    
    
class CT_FileVersion(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    appName = OptionalAttribute('appName', XsdString) # Application Name
    lastEdited = OptionalAttribute('lastEdited', XsdString) # Last Edited Version
    lowestEdited = OptionalAttribute('lowestEdited', XsdString) # Lowest Edited Version
    rupBuild = OptionalAttribute('rupBuild', XsdString) # Build Version
    codeName = OptionalAttribute('codeName', ST_Guid) # Code Name
    
class CT_BookViews(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    workbookView = OneOrMore('ssml:workbookView') # Workbook View
    
    
class CT_BookView(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    extLst = ZeroOrOne('ssml:extLst') # None
    
    visibility = OptionalAttribute('visibility', ST_Visibility, 'visible') # Visibility
    minimized = OptionalAttribute('minimized', XsdBoolean, False) # Minimized
    showHorizontalScroll = OptionalAttribute('showHorizontalScroll', XsdBoolean, True) # Show Horizontal Scroll
    showVerticalScroll = OptionalAttribute('showVerticalScroll', XsdBoolean, True) # Show Vertical Scroll
    showSheetTabs = OptionalAttribute('showSheetTabs', XsdBoolean, True) # Show Sheet Tabs
    xWindow = OptionalAttribute('xWindow', XsdInt) # Upper Left Corner (X Coordinate)
    yWindow = OptionalAttribute('yWindow', XsdInt) # Upper Left Corner (Y Coordinate)
    windowWidth = OptionalAttribute('windowWidth', XsdUnsignedInt) # Window Width
    windowHeight = OptionalAttribute('windowHeight', XsdUnsignedInt) # Window Height
    tabRatio = OptionalAttribute('tabRatio', XsdUnsignedInt, 600) # Sheet Tab Ratio
    firstSheet = OptionalAttribute('firstSheet', XsdUnsignedInt, 0) # First Sheet
    activeTab = OptionalAttribute('activeTab', XsdUnsignedInt, 0) # Active Sheet Index
    autoFilterDateGrouping = OptionalAttribute('autoFilterDateGrouping', XsdBoolean, True) # AutoFilter Date Grouping
    
class CT_CustomWorkbookViews(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    customWorkbookView = OneOrMore('ssml:customWorkbookView') # Custom Workbook View
    
    
class CT_CustomWorkbookView(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    extLst = ZeroOrMore('ssml:extLst') # None
    
    name = RequiredAttribute('name', ST_Xstring) # Custom View Name
    guid = RequiredAttribute('guid', ST_Guid) # Custom View GUID
    autoUpdate = OptionalAttribute('autoUpdate', XsdBoolean, False) # Auto Update
    mergeInterval = OptionalAttribute('mergeInterval', XsdUnsignedInt) # Merge Interval
    changesSavedWin = OptionalAttribute('changesSavedWin', XsdBoolean, False) # Changes Saved Win
    onlySync = OptionalAttribute('onlySync', XsdBoolean, False) # Only Synch
    personalView = OptionalAttribute('personalView', XsdBoolean, False) # Personal View
    includePrintSettings = OptionalAttribute('includePrintSettings', XsdBoolean, True) # Include Print Settings
    includeHiddenRowCol = OptionalAttribute('includeHiddenRowCol', XsdBoolean, True) # Include Hidden Rows & Columns
    maximized = OptionalAttribute('maximized', XsdBoolean, False) # Maximized
    minimized = OptionalAttribute('minimized', XsdBoolean, False) # Minimized
    showHorizontalScroll = OptionalAttribute('showHorizontalScroll', XsdBoolean, True) # Show Horizontal Scroll
    showVerticalScroll = OptionalAttribute('showVerticalScroll', XsdBoolean, True) # Show Vertical Scroll
    showSheetTabs = OptionalAttribute('showSheetTabs', XsdBoolean, True) # Show Sheet Tabs
    xWindow = OptionalAttribute('xWindow', XsdInt, 0) # Top Left Corner (X Coordinate)
    yWindow = OptionalAttribute('yWindow', XsdInt, 0) # Top Left Corner (Y Coordinate)
    windowWidth = RequiredAttribute('windowWidth', XsdUnsignedInt) # Window Width
    windowHeight = RequiredAttribute('windowHeight', XsdUnsignedInt) # Window Height
    tabRatio = OptionalAttribute('tabRatio', XsdUnsignedInt, 600) # Sheet Tab Ratio
    activeSheetId = RequiredAttribute('activeSheetId', XsdUnsignedInt) # Active Sheet in Book View
    showFormulaBar = OptionalAttribute('showFormulaBar', XsdBoolean, True) # Show Formula Bar
    showStatusbar = OptionalAttribute('showStatusbar', XsdBoolean, True) # Show Status Bar
    showComments = OptionalAttribute('showComments', ST_Comments, 'commIndicator') # Show Comments
    showObjects = OptionalAttribute('showObjects', ST_Objects, 'all') # Show Objects
    
class CT_Sheets(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    sheet = OneOrMore('ssml:sheet') # Sheet Information
    
    
class CT_Sheet(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    name = RequiredAttribute('name', ST_Xstring) # Sheet Name
    sheetId = RequiredAttribute('sheetId', XsdUnsignedInt) # Sheet Tab Id
    state = OptionalAttribute('state', ST_SheetState, 'visible') # Visible State
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_WorkbookPr(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    date1904 = OptionalAttribute('date1904', XsdBoolean, False) # Date 1904
    showObjects = OptionalAttribute('showObjects', ST_Objects, 'all') # Show Objects
    showBorderUnselectedTables = OptionalAttribute('showBorderUnselectedTables', XsdBoolean, True) # Show Border Unselected Table
    filterPrivacy = OptionalAttribute('filterPrivacy', XsdBoolean, False) # Filter Privacy
    promptedSolutions = OptionalAttribute('promptedSolutions', XsdBoolean, False) # Prompted Solutions
    showInkAnnotation = OptionalAttribute('showInkAnnotation', XsdBoolean, True) # Show Ink Annotations
    backupFile = OptionalAttribute('backupFile', XsdBoolean, False) # Create Backup File
    saveExternalLinkValues = OptionalAttribute('saveExternalLinkValues', XsdBoolean, True) # Save External Link Values
    updateLinks = OptionalAttribute('updateLinks', ST_UpdateLinks, 'userSet') # Update Links Behavior
    codeName = OptionalAttribute('codeName', XsdString) # Code Name
    hidePivotFieldList = OptionalAttribute('hidePivotFieldList', XsdBoolean, False) # Hide Pivot Field List
    showPivotChartFilter = OptionalAttribute('showPivotChartFilter', XsdBoolean) # Show Pivot Chart Filter
    allowRefreshQuery = OptionalAttribute('allowRefreshQuery', XsdBoolean, False) # Allow Refresh Query
    publishItems = OptionalAttribute('publishItems', XsdBoolean, False) # Publish Items
    checkCompatibility = OptionalAttribute('checkCompatibility', XsdBoolean, False) # Check Compatibility On Save
    autoCompressPictures = OptionalAttribute('autoCompressPictures', XsdBoolean, True) # Auto Compress Pictures
    refreshAllConnections = OptionalAttribute('refreshAllConnections', XsdBoolean, False) # Refresh all Connections on Open
    defaultThemeVersion = OptionalAttribute('defaultThemeVersion', XsdUnsignedInt) # Default Theme Version
    
class CT_SmartTagPr(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    embed = OptionalAttribute('embed', XsdBoolean, False) # Embed SmartTags
    show = OptionalAttribute('show', ST_SmartTagShow, 'all') # Show Smart Tags
    
class CT_SmartTagTypes(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    smartTagType = ZeroOrMore('ssml:smartTagType') # Smart Tag Type
    
    
class CT_SmartTagType(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    namespaceUri = OptionalAttribute('namespaceUri', ST_Xstring) # SmartTag Namespace URI
    name = OptionalAttribute('name', ST_Xstring) # Name
    url = OptionalAttribute('url', ST_Xstring) # Smart Tag URL
    
class CT_FileRecoveryPr(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    autoRecover = OptionalAttribute('autoRecover', XsdBoolean, True) # Auto Recover
    crashSave = OptionalAttribute('crashSave', XsdBoolean, False) # Crash Save
    dataExtractLoad = OptionalAttribute('dataExtractLoad', XsdBoolean, False) # Data Extract Load
    repairLoad = OptionalAttribute('repairLoad', XsdBoolean, False) # Repair Load
    
class CT_CalcPr(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    calcId = OptionalAttribute('calcId', XsdUnsignedInt) # Calculation Id
    calcMode = OptionalAttribute('calcMode', ST_CalcMode, 'auto') # Calculation Mode
    fullCalcOnLoad = OptionalAttribute('fullCalcOnLoad', XsdBoolean, False) # Full Calculation On Load
    refMode = OptionalAttribute('refMode', ST_RefMode, 'A1') # Reference Mode
    iterate = OptionalAttribute('iterate', XsdBoolean, False) # Calculation Iteration
    iterateCount = OptionalAttribute('iterateCount', XsdUnsignedInt, 100) # Iteration Count
    iterateDelta = OptionalAttribute('iterateDelta', XsdDouble, 0.001) # Iterative Calculation Delta
    fullPrecision = OptionalAttribute('fullPrecision', XsdBoolean, True) # Full Precision Calculation
    calcCompleted = OptionalAttribute('calcCompleted', XsdBoolean, True) # Calc Completed
    calcOnSave = OptionalAttribute('calcOnSave', XsdBoolean, True) # Calculate On Save
    concurrentCalc = OptionalAttribute('concurrentCalc', XsdBoolean, True) # Concurrent Calculations
    concurrentManualCount = OptionalAttribute('concurrentManualCount', XsdUnsignedInt) # Concurrent Thread Manual Count
    forceFullCalc = OptionalAttribute('forceFullCalc', XsdBoolean) # Force Full Calculation
    
class CT_DefinedNames(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    definedName = ZeroOrMore('ssml:definedName') # Defined Name
    
    
class CT_DefinedName(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    # simpleContent[]: 
      
    
class CT_ExternalReferences(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    externalReference = OneOrMore('ssml:externalReference') # External Reference
    
    
class CT_ExternalReference(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_SheetBackgroundPicture(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_PivotCaches(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    pivotCache = OneOrMore('ssml:pivotCache') # PivotCache
    
    
class CT_PivotCache(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    cacheId = RequiredAttribute('cacheId', XsdUnsignedInt) # PivotCache Id
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_FileSharing(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    readOnlyRecommended = OptionalAttribute('readOnlyRecommended', XsdBoolean, False) # Read Only Recommended
    userName = OptionalAttribute('userName', ST_Xstring) # User Name
    reservationPassword = OptionalAttribute('reservationPassword', ST_UnsignedShortHex) # Write Reservation Password
    
class CT_OleSize(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    ref = RequiredAttribute('ref', ST_Ref) # Reference
    
class CT_WorkbookProtection(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    workbookPassword = OptionalAttribute('workbookPassword', ST_UnsignedShortHex) # Workbook Password
    revisionsPassword = OptionalAttribute('revisionsPassword', ST_UnsignedShortHex) # Revisions Password
    lockStructure = OptionalAttribute('lockStructure', XsdBoolean, False) # Lock Structure
    lockWindows = OptionalAttribute('lockWindows', XsdBoolean, False) # Lock Windows
    lockRevision = OptionalAttribute('lockRevision', XsdBoolean, False) # Lock Revisions
    
class CT_WebPublishing(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    css = OptionalAttribute('css', XsdBoolean, True) # Use CSS
    thicket = OptionalAttribute('thicket', XsdBoolean, True) # Thicket
    longFileNames = OptionalAttribute('longFileNames', XsdBoolean, True) # Enable Long File Names
    vml = OptionalAttribute('vml', XsdBoolean, False) # VML in Browsers
    allowPng = OptionalAttribute('allowPng', XsdBoolean, False) # Allow PNG
    targetScreenSize = OptionalAttribute('targetScreenSize', ST_TargetScreenSize, '800x600') # Target Screen Size
    dpi = OptionalAttribute('dpi', XsdUnsignedInt, 96) # DPI
    codePage = OptionalAttribute('codePage', XsdUnsignedInt) # Code Page
    
class CT_FunctionGroups(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    functionGroup = ZeroOrMore('ssml:functionGroup') # Function Group
    
    builtInGroupCount = OptionalAttribute('builtInGroupCount', XsdUnsignedInt, 16) # Built-in Function Group Count
    
class CT_FunctionGroup(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    name = OptionalAttribute('name', ST_Xstring) # Name
    
class CT_WebPublishObjects(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    webPublishObject = OneOrMore('ssml:webPublishObject') # Web Publishing Object
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Count
    
class CT_WebPublishObject(BaseOxmlElement):
    """
    Complex type (sml-workbook.xsd)
    
    """
    id = RequiredAttribute('id', XsdUnsignedInt) # Id
    divId = RequiredAttribute('divId', ST_Xstring) # Div Id
    sourceObject = OptionalAttribute('sourceObject', ST_Xstring) # Source Object
    destinationFile = RequiredAttribute('destinationFile', ST_Xstring) # Destination File
    title = OptionalAttribute('title', ST_Xstring) # Title
    autoRepublish = OptionalAttribute('autoRepublish', XsdBoolean, False) # Auto Republish
    
    
    
