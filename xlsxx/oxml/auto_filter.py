
from docxx.oxml.xmlchemy import (
    BaseOxmlElement, ZeroOrOne, ZeroOrMore, OneOrMore, OneAndOnlyOne,
    RequiredAttribute, OptionalAttribute, 
)
from docxx.oxml.simpletypes import (
    XsdBoolean, XsdDouble, XsdUnsignedInt, XsdUnsignedShort,
)
from xlsxx.oxml.simpletypes import (
    ST_CalendarType, ST_DateTimeGrouping, ST_DynamicFilterType, 
    ST_FilterOperator, ST_IconSetType, 
    ST_SortBy, ST_SortMethod, ST_Ref, ST_Xstring, ST_DxfId,
)


class CT_AutoFilter(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    filterColumn = ZeroOrMore('ssml:filterColumn') # AutoFilter Column
    sortState = ZeroOrOne('ssml:sortState') # Sort State for Auto Filter
    extLst = ZeroOrOne('ssml:extLst') # None
    
    ref = OptionalAttribute('ref', ST_Ref) # Cell or Range Reference
    
class CT_FilterColumn(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    filters = ZeroOrOne('ssml:filters') # Filter Criteria
    top10 = ZeroOrOne('ssml:top10') # Top 10
    customFilters = ZeroOrOne('ssml:customFilters') # Custom Filters
    dynamicFilter = ZeroOrOne('ssml:dynamicFilter') # Dynamic Filter
    colorFilter = ZeroOrOne('ssml:colorFilter') # Color Filter Criteria
    iconFilter = ZeroOrOne('ssml:iconFilter') # Icon Filter
    extLst = ZeroOrOne('ssml:extLst') # None
    
    colId = RequiredAttribute('colId', XsdUnsignedInt) # Filter Column Data
    hiddenButton = OptionalAttribute('hiddenButton', XsdBoolean, False) # Hidden AutoFilter Button
    showButton = OptionalAttribute('showButton', XsdBoolean, True) # Show Filter Button
    
class CT_Filters(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    filter = ZeroOrMore('ssml:filter') # Filter
    dateGroupItem = ZeroOrMore('ssml:dateGroupItem') # Date Grouping
    
    blank = OptionalAttribute('blank', XsdBoolean, False) # Filter by Blank
    calendarType = OptionalAttribute('calendarType', ST_CalendarType, 'none') # Calendar Type
    
class CT_Filter(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    val = OptionalAttribute('val', ST_Xstring) # Filter Value
    
class CT_CustomFilters(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    customFilter = ZeroOrMore('ssml:customFilter') # Custom Filter Criteria
    
    and_ = OptionalAttribute('and', XsdBoolean, False) # And
    
class CT_CustomFilter(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    operator = OptionalAttribute('operator', ST_FilterOperator, 'equal') # Filter Comparison Operator
    val = OptionalAttribute('val', ST_Xstring) # Top or Bottom Value
    
class CT_Top10(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    top = OptionalAttribute('top', XsdBoolean, True) # Top
    percent = OptionalAttribute('percent', XsdBoolean, False) # Filter by Percent
    val = RequiredAttribute('val', XsdDouble) # Top or Bottom Value
    filterVal = OptionalAttribute('filterVal', XsdDouble) # Filter Value
    
class CT_ColorFilter(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    dxfId = OptionalAttribute('dxfId', ST_DxfId) # Differential Format Record Id
    cellColor = OptionalAttribute('cellColor', XsdBoolean, True) # Filter By Cell Color
    
class CT_IconFilter(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    iconSet = RequiredAttribute('iconSet', ST_IconSetType) # Icon Set
    iconId = OptionalAttribute('iconId', XsdUnsignedInt) # Icon Id
    
class CT_DynamicFilter(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    type = RequiredAttribute('type', ST_DynamicFilterType) # Dynamic filter type
    val = OptionalAttribute('val', XsdDouble) # Value
    maxVal = OptionalAttribute('maxVal', XsdDouble) # Max Value
    
class CT_SortState(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    sortCondition = ZeroOrMore('ssml:sortCondition') # Sort Condition
    extLst = ZeroOrOne('ssml:extLst') # None
    
    columnSort = OptionalAttribute('columnSort', XsdBoolean, False) # Sort by Columns
    caseSensitive = OptionalAttribute('caseSensitive', XsdBoolean, False) # Case Sensitive
    sortMethod = OptionalAttribute('sortMethod', ST_SortMethod, 'none') # Sort Method
    ref = RequiredAttribute('ref', ST_Ref) # Sort Range
    
class CT_SortCondition(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    descending = OptionalAttribute('descending', XsdBoolean, False) # Descending
    sortBy = OptionalAttribute('sortBy', ST_SortBy, 'value') # Sort By
    ref = RequiredAttribute('ref', ST_Ref) # Reference
    customList = OptionalAttribute('customList', ST_Xstring) # Custom List
    dxfId = OptionalAttribute('dxfId', ST_DxfId) # Format Id
    iconSet = OptionalAttribute('iconSet', ST_IconSetType, '3Arrows') # Icon Set
    iconId = OptionalAttribute('iconId', XsdUnsignedInt) # Icon Id
    
class CT_DateGroupItem(BaseOxmlElement):
    """
    Complex type (sml-autoFilter.xsd)
    
    """
    year = RequiredAttribute('year', XsdUnsignedShort) # Year
    month = OptionalAttribute('month', XsdUnsignedShort) # Month
    day = OptionalAttribute('day', XsdUnsignedShort) # Day
    hour = OptionalAttribute('hour', XsdUnsignedShort) # Hour
    minute = OptionalAttribute('minute', XsdUnsignedShort) # Minute
    second = OptionalAttribute('second', XsdUnsignedShort) # Second
    dateTimeGrouping = RequiredAttribute('dateTimeGrouping', ST_DateTimeGrouping) # Date Time Grouping
    
