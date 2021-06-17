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
    Complex type: 
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
    Complex type: 
    """
    appName = OptionalAttribute('ssml:appName', XsdString) # Application Name
    lastEdited = OptionalAttribute('ssml:lastEdited', XsdString) # Last Edited Version
    lowestEdited = OptionalAttribute('ssml:lowestEdited', XsdString) # Lowest Edited Version
    rupBuild = OptionalAttribute('ssml:rupBuild', XsdString) # Build Version
    codeName = OptionalAttribute('ssml:codeName', ST_Guid) # Code Name
    
class CT_BookViews(BaseOxmlElement):
    """
    Complex type: 
    """
    workbookView = OneOrMore('ssml:workbookView') # Workbook View
    
    
class CT_BookView(BaseOxmlElement):
    """
    Complex type: 
    """
    extLst = ZeroOrOne('ssml:extLst') # None
    
    visibility = OptionalAttribute('ssml:visibility', ST_Visibility) # Visibility
    minimized = OptionalAttribute('ssml:minimized', XsdBoolean) # Minimized
    showHorizontalScroll = OptionalAttribute('ssml:showHorizontalScroll', XsdBoolean) # Show Horizontal Scroll
    showVerticalScroll = OptionalAttribute('ssml:showVerticalScroll', XsdBoolean) # Show Vertical Scroll
    showSheetTabs = OptionalAttribute('ssml:showSheetTabs', XsdBoolean) # Show Sheet Tabs
    xWindow = OptionalAttribute('ssml:xWindow', XsdInt) # Upper Left Corner (X Coordinate)
    yWindow = OptionalAttribute('ssml:yWindow', XsdInt) # Upper Left Corner (Y Coordinate)
    windowWidth = OptionalAttribute('ssml:windowWidth', XsdUnsignedInt) # Window Width
    windowHeight = OptionalAttribute('ssml:windowHeight', XsdUnsignedInt) # Window Height
    tabRatio = OptionalAttribute('ssml:tabRatio', XsdUnsignedInt) # Sheet Tab Ratio
    firstSheet = OptionalAttribute('ssml:firstSheet', XsdUnsignedInt) # First Sheet
    activeTab = OptionalAttribute('ssml:activeTab', XsdUnsignedInt) # Active Sheet Index
    autoFilterDateGrouping = OptionalAttribute('ssml:autoFilterDateGrouping', XsdBoolean) # AutoFilter Date Grouping
    
class CT_CustomWorkbookViews(BaseOxmlElement):
    """
    Complex type: 
    """
    customWorkbookView = OneOrMore('ssml:customWorkbookView') # Custom Workbook View
    
    
class CT_CustomWorkbookView(BaseOxmlElement):
    """
    Complex type: 
    """
    extLst = ZeroOrOne('ssml:extLst') # None
    
    name = RequiredAttribute('ssml:name', ST_Xstring) # Custom View Name
    guid = RequiredAttribute('ssml:guid', ST_Guid) # Custom View GUID
    autoUpdate = OptionalAttribute('ssml:autoUpdate', XsdBoolean) # Auto Update
    mergeInterval = OptionalAttribute('ssml:mergeInterval', XsdUnsignedInt) # Merge Interval
    changesSavedWin = OptionalAttribute('ssml:changesSavedWin', XsdBoolean) # Changes Saved Win
    onlySync = OptionalAttribute('ssml:onlySync', XsdBoolean) # Only Synch
    personalView = OptionalAttribute('ssml:personalView', XsdBoolean) # Personal View
    includePrintSettings = OptionalAttribute('ssml:includePrintSettings', XsdBoolean) # Include Print Settings
    includeHiddenRowCol = OptionalAttribute('ssml:includeHiddenRowCol', XsdBoolean) # Include Hidden Rows & Columns
    maximized = OptionalAttribute('ssml:maximized', XsdBoolean) # Maximized
    minimized = OptionalAttribute('ssml:minimized', XsdBoolean) # Minimized
    showHorizontalScroll = OptionalAttribute('ssml:showHorizontalScroll', XsdBoolean) # Show Horizontal Scroll
    showVerticalScroll = OptionalAttribute('ssml:showVerticalScroll', XsdBoolean) # Show Vertical Scroll
    showSheetTabs = OptionalAttribute('ssml:showSheetTabs', XsdBoolean) # Show Sheet Tabs
    xWindow = OptionalAttribute('ssml:xWindow', XsdInt) # Top Left Corner (X Coordinate)
    yWindow = OptionalAttribute('ssml:yWindow', XsdInt) # Top Left Corner (Y Coordinate)
    windowWidth = RequiredAttribute('ssml:windowWidth', XsdUnsignedInt) # Window Width
    windowHeight = RequiredAttribute('ssml:windowHeight', XsdUnsignedInt) # Window Height
    tabRatio = OptionalAttribute('ssml:tabRatio', XsdUnsignedInt) # Sheet Tab Ratio
    activeSheetId = RequiredAttribute('ssml:activeSheetId', XsdUnsignedInt) # Active Sheet in Book View
    showFormulaBar = OptionalAttribute('ssml:showFormulaBar', XsdBoolean) # Show Formula Bar
    showStatusbar = OptionalAttribute('ssml:showStatusbar', XsdBoolean) # Show Status Bar
    showComments = OptionalAttribute('ssml:showComments', ST_Comments) # Show Comments
    showObjects = OptionalAttribute('ssml:showObjects', ST_Objects) # Show Objects
    
class CT_Sheets(BaseOxmlElement):
    """
    Complex type: 
    """
    sheet = OneOrMore('ssml:sheet') # Sheet Information
    
    
class CT_Sheet(BaseOxmlElement):
    """
    Complex type: 
    """
    name = RequiredAttribute('ssml:name', ST_Xstring) # Sheet Name
    id = RequiredAttribute('ssml:sheetId', XsdUnsignedInt) # Sheet Tab Id
    state = OptionalAttribute('ssml:state', ST_SheetState) # Visible State
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_WorkbookPr(BaseOxmlElement):
    """
    Complex type: 
    """
    date1904 = OptionalAttribute('ssml:date1904', XsdBoolean) # Date 1904
    showObjects = OptionalAttribute('ssml:showObjects', ST_Objects) # Show Objects
    showBorderUnselectedTables = OptionalAttribute('ssml:showBorderUnselectedTables', XsdBoolean) # Show Border Unselected Table
    filterPrivacy = OptionalAttribute('ssml:filterPrivacy', XsdBoolean) # Filter Privacy
    promptedSolutions = OptionalAttribute('ssml:promptedSolutions', XsdBoolean) # Prompted Solutions
    showInkAnnotation = OptionalAttribute('ssml:showInkAnnotation', XsdBoolean) # Show Ink Annotations
    backupFile = OptionalAttribute('ssml:backupFile', XsdBoolean) # Create Backup File
    saveExternalLinkValues = OptionalAttribute('ssml:saveExternalLinkValues', XsdBoolean) # Save External Link Values
    updateLinks = OptionalAttribute('ssml:updateLinks', ST_UpdateLinks) # Update Links Behavior
    codeName = OptionalAttribute('ssml:codeName', XsdString) # Code Name
    hidePivotFieldList = OptionalAttribute('ssml:hidePivotFieldList', XsdBoolean) # Hide Pivot Field List
    showPivotChartFilter = OptionalAttribute('ssml:showPivotChartFilter', XsdBoolean) # Show Pivot Chart Filter
    allowRefreshQuery = OptionalAttribute('ssml:allowRefreshQuery', XsdBoolean) # Allow Refresh Query
    publishItems = OptionalAttribute('ssml:publishItems', XsdBoolean) # Publish Items
    checkCompatibility = OptionalAttribute('ssml:checkCompatibility', XsdBoolean) # Check Compatibility On Save
    autoCompressPictures = OptionalAttribute('ssml:autoCompressPictures', XsdBoolean) # Auto Compress Pictures
    refreshAllConnections = OptionalAttribute('ssml:refreshAllConnections', XsdBoolean) # Refresh all Connections on Open
    defaultThemeVersion = OptionalAttribute('ssml:defaultThemeVersion', XsdUnsignedInt) # Default Theme Version
    
class CT_SmartTagPr(BaseOxmlElement):
    """
    Complex type: 
    """
    embed = OptionalAttribute('ssml:embed', XsdBoolean) # Embed SmartTags
    show = OptionalAttribute('ssml:show', ST_SmartTagShow) # Show Smart Tags
    
class CT_SmartTagTypes(BaseOxmlElement):
    """
    Complex type: 
    """
    smartTagType = ZeroOrMore('ssml:smartTagType') # Smart Tag Type
    
    
class CT_SmartTagType(BaseOxmlElement):
    """
    Complex type: 
    """
    namespaceUri = OptionalAttribute('ssml:namespaceUri', ST_Xstring) # SmartTag Namespace URI
    name = OptionalAttribute('ssml:name', ST_Xstring) # Name
    url = OptionalAttribute('ssml:url', ST_Xstring) # Smart Tag URL
    
class CT_FileRecoveryPr(BaseOxmlElement):
    """
    Complex type: 
    """
    autoRecover = OptionalAttribute('ssml:autoRecover', XsdBoolean) # Auto Recover
    crashSave = OptionalAttribute('ssml:crashSave', XsdBoolean) # Crash Save
    dataExtractLoad = OptionalAttribute('ssml:dataExtractLoad', XsdBoolean) # Data Extract Load
    repairLoad = OptionalAttribute('ssml:repairLoad', XsdBoolean) # Repair Load
    
class CT_CalcPr(BaseOxmlElement):
    """
    Complex type: 
    """
    calcId = OptionalAttribute('ssml:calcId', XsdUnsignedInt) # Calculation Id
    calcMode = OptionalAttribute('ssml:calcMode', ST_CalcMode) # Calculation Mode
    fullCalcOnLoad = OptionalAttribute('ssml:fullCalcOnLoad', XsdBoolean) # Full Calculation On Load
    refMode = OptionalAttribute('ssml:refMode', ST_RefMode) # Reference Mode
    iterate = OptionalAttribute('ssml:iterate', XsdBoolean) # Calculation Iteration
    iterateCount = OptionalAttribute('ssml:iterateCount', XsdUnsignedInt) # Iteration Count
    iterateDelta = OptionalAttribute('ssml:iterateDelta', XsdDouble) # Iterative Calculation Delta
    fullPrecision = OptionalAttribute('ssml:fullPrecision', XsdBoolean) # Full Precision Calculation
    calcCompleted = OptionalAttribute('ssml:calcCompleted', XsdBoolean) # Calc Completed
    calcOnSave = OptionalAttribute('ssml:calcOnSave', XsdBoolean) # Calculate On Save
    concurrentCalc = OptionalAttribute('ssml:concurrentCalc', XsdBoolean) # Concurrent Calculations
    concurrentManualCount = OptionalAttribute('ssml:concurrentManualCount', XsdUnsignedInt) # Concurrent Thread Manual Count
    forceFullCalc = OptionalAttribute('ssml:forceFullCalc', XsdBoolean) # Force Full Calculation
    
class CT_DefinedNames(BaseOxmlElement):
    """
    Complex type: 
    """
    definedName = ZeroOrMore('ssml:definedName') # Defined Name
    
    
class CT_DefinedName(BaseOxmlElement):
    """
    Complex type: 
    """
    # simpleContent[]: 
      
    
class CT_ExternalReferences(BaseOxmlElement):
    """
    Complex type: 
    """
    externalReference = OneOrMore('ssml:externalReference') # External Reference
    
    
class CT_ExternalReference(BaseOxmlElement):
    """
    Complex type: 
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_SheetBackgroundPicture(BaseOxmlElement):
    """
    Complex type: 
    """
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_PivotCaches(BaseOxmlElement):
    """
    Complex type: 
    """
    pivotCache = OneOrMore('ssml:pivotCache') # PivotCache
    
    
class CT_PivotCache(BaseOxmlElement):
    """
    Complex type: 
    """
    cacheId = RequiredAttribute('ssml:cacheId', XsdUnsignedInt) # PivotCache Id
    rid = RequiredAttribute('r:id', ST_RelationshipId) # Relationship Id
    
class CT_FileSharing(BaseOxmlElement):
    """
    Complex type: 
    """
    readOnlyRecommended = OptionalAttribute('ssml:readOnlyRecommended', XsdBoolean) # Read Only Recommended
    userName = OptionalAttribute('ssml:userName', ST_Xstring) # User Name
    reservationPassword = OptionalAttribute('ssml:reservationPassword', ST_UnsignedShortHex) # Write Reservation Password
    
class CT_OleSize(BaseOxmlElement):
    """
    Complex type: 
    """
    ref = RequiredAttribute('ssml:ref', ST_Ref) # Reference
    
class CT_WorkbookProtection(BaseOxmlElement):
    """
    Complex type: 
    """
    workbookPassword = OptionalAttribute('ssml:workbookPassword', ST_UnsignedShortHex) # Workbook Password
    revisionsPassword = OptionalAttribute('ssml:revisionsPassword', ST_UnsignedShortHex) # Revisions Password
    lockStructure = OptionalAttribute('ssml:lockStructure', XsdBoolean) # Lock Structure
    lockWindows = OptionalAttribute('ssml:lockWindows', XsdBoolean) # Lock Windows
    lockRevision = OptionalAttribute('ssml:lockRevision', XsdBoolean) # Lock Revisions
    
class CT_WebPublishing(BaseOxmlElement):
    """
    Complex type: 
    """
    css = OptionalAttribute('ssml:css', XsdBoolean) # Use CSS
    thicket = OptionalAttribute('ssml:thicket', XsdBoolean) # Thicket
    longFileNames = OptionalAttribute('ssml:longFileNames', XsdBoolean) # Enable Long File Names
    vml = OptionalAttribute('ssml:vml', XsdBoolean) # VML in Browsers
    allowPng = OptionalAttribute('ssml:allowPng', XsdBoolean) # Allow PNG
    targetScreenSize = OptionalAttribute('ssml:targetScreenSize', ST_TargetScreenSize) # Target Screen Size
    dpi = OptionalAttribute('ssml:dpi', XsdUnsignedInt) # DPI
    codePage = OptionalAttribute('ssml:codePage', XsdUnsignedInt) # Code Page
    
class CT_FunctionGroups(BaseOxmlElement):
    """
    Complex type: 
    """
    functionGroup = ZeroOrOne('ssml:functionGroup') # Function Group
    
    builtInGroupCount = OptionalAttribute('ssml:builtInGroupCount', XsdUnsignedInt) # Built-in Function Group Count
    
class CT_FunctionGroup(BaseOxmlElement):
    """
    Complex type: 
    """
    name = OptionalAttribute('ssml:name', ST_Xstring) # Name
    
class CT_WebPublishObjects(BaseOxmlElement):
    """
    Complex type: 
    """
    webPublishObject = OneOrMore('ssml:webPublishObject') # Web Publishing Object
    
    count = OptionalAttribute('ssml:count', XsdUnsignedInt) # Count
    
class CT_WebPublishObject(BaseOxmlElement):
    """
    Complex type: 
    """
    id = RequiredAttribute('ssml:id', XsdUnsignedInt) # Id
    divId = RequiredAttribute('ssml:divId', ST_Xstring) # Div Id
    sourceObject = OptionalAttribute('ssml:sourceObject', ST_Xstring) # Source Object
    destinationFile = RequiredAttribute('ssml:destinationFile', ST_Xstring) # Destination File
    title = OptionalAttribute('ssml:title', ST_Xstring) # Title
    autoRepublish = OptionalAttribute('ssml:autoRepublish', XsdBoolean) # Auto Republish
    
    
    
