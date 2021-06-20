
from docxx.oxml.xmlchemy import (
    BaseOxmlElement, ZeroOrMore, 
    RequiredAttribute, OptionalAttribute, 
)
from docxx.oxml.simpletypes import (
    XsdBoolean, XsdInt, XsdUnsignedInt, XsdUnsignedShort,
)
from xlsxx.oxml.simpletypes import (
    ST_Axis, ST_PivotAreaType, ST_Ref
)



class CT_PivotArea(BaseOxmlElement):
    """
    Complex type (sml-pivotTableShared.xsd)
    
    """
    references = ZeroOrMore('ssml:references') # References
    extLst = ZeroOrMore('ssml:extLst') # Future Feature Data Storage Area
    
    field = OptionalAttribute('field', XsdInt) # Field Index
    type = OptionalAttribute('type', ST_PivotAreaType) # Rule Type
    dataOnly = OptionalAttribute('dataOnly', XsdBoolean) # Data Only
    labelOnly = OptionalAttribute('labelOnly', XsdBoolean) # Labels Only
    grandRow = OptionalAttribute('grandRow', XsdBoolean) # Include Row Grand Total
    grandCol = OptionalAttribute('grandCol', XsdBoolean) # Include Column Grand Total
    cacheIndex = OptionalAttribute('cacheIndex', XsdBoolean) # Cache Index
    outline = OptionalAttribute('outline', XsdBoolean) # Outline
    offset = OptionalAttribute('offset', ST_Ref) # Offset Reference
    collapsedLevelsAreSubtotals = OptionalAttribute('collapsedLevelsAreSubtotals', XsdBoolean) # Collapsed Levels Are Subtotals
    axis = OptionalAttribute('axis', ST_Axis) # Axis
    fieldPosition = OptionalAttribute('fieldPosition', XsdUnsignedInt) # Field Position
    
class CT_PivotAreaReferences(BaseOxmlElement):
    """
    Complex type (sml-pivotTableShared.xsd)
    
    """
    reference = ZeroOrMore('ssml:reference') # Reference
    
    count = OptionalAttribute('count', XsdUnsignedInt) # Pivot Filter Count
    
class CT_PivotAreaReference(BaseOxmlElement):
    """
    Complex type (sml-pivotTableShared.xsd)
    
    """
    x = ZeroOrMore('ssml:x') # Field Item
    extLst = ZeroOrMore('ssml:extLst') # None
    
    field = OptionalAttribute('field', XsdUnsignedInt) # Field Index
    count = OptionalAttribute('count', XsdUnsignedInt) # Item Index Count
    selected = OptionalAttribute('selected', XsdBoolean) # Selected
    byPosition = OptionalAttribute('byPosition', XsdBoolean) # Positional Reference
    relative = OptionalAttribute('relative', XsdBoolean) # Relative Reference
    defaultSubtotal = OptionalAttribute('defaultSubtotal', XsdBoolean) # Include Default Filter
    sumSubtotal = OptionalAttribute('sumSubtotal', XsdBoolean) # Include Sum Filter
    countASubtotal = OptionalAttribute('countASubtotal', XsdBoolean) # Include CountA Filter
    avgSubtotal = OptionalAttribute('avgSubtotal', XsdBoolean) # Include Average Filter
    maxSubtotal = OptionalAttribute('maxSubtotal', XsdBoolean) # Include Maximum Filter
    minSubtotal = OptionalAttribute('minSubtotal', XsdBoolean) # Include Minimum Filter
    productSubtotal = OptionalAttribute('productSubtotal', XsdBoolean) # Include Product Filter
    countSubtotal = OptionalAttribute('countSubtotal', XsdBoolean) # Include Count Subtotal
    stdDevSubtotal = OptionalAttribute('stdDevSubtotal', XsdBoolean) # Include StdDev Filter
    stdDevPSubtotal = OptionalAttribute('stdDevPSubtotal', XsdBoolean) # Include StdDevP Filter
    varSubtotal = OptionalAttribute('varSubtotal', XsdBoolean) # Include Var Filter
    varPSubtotal = OptionalAttribute('varPSubtotal', XsdBoolean) # Include VarP Filter
    
class CT_Index(BaseOxmlElement):
    """
    Complex type (sml-pivotTableShared.xsd)
    
    """
    v = RequiredAttribute('v', XsdUnsignedInt) # Shared Items Index
