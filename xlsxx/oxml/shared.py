from docxx.oxml.xmlchemy import (
    BaseOxmlElement, ZeroOrOne, ZeroOrMore, OneOrMore, OneAndOnlyOne,
    RequiredAttribute, OptionalAttribute, 
)
from docxx.oxml.simpletypes import (
    ST_DecimalNumber, ST_OnOff, ST_String, XsdToken
)
from xlsxx.oxml.simpletypes import ST_Xstring

#
# sml basic elements
#
class CT_Xstring(BaseOxmlElement):
    pass
    
class CT_XStringElement(BaseOxmlElement):
    v = RequiredAttribute('v', ST_Xstring) # Value

class CT_Extension(BaseOxmlElement):
    # any[processContents='lax']: None
    uri = OptionalAttribute('uri', XsdToken) # URI*

class CT_ExtensionList(BaseOxmlElement):
    ext = ZeroOrMore("ssml:ext")
    