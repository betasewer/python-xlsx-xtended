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
    XsdUnsignedInt, ST_RelationshipId, XsdToken
)
from xlsxx.oxml.simpletypes import ST_Xstring

class CT_Workbook(BaseOxmlElement):
    sheets = OneAndOnlyOne('ssml:sheets')

class CT_Sheets(BaseOxmlElement):
    sheet = OneOrMore('ssml:sheet')

class CT_Sheet(BaseOxmlElement):
    name = RequiredAttribute("name", ST_Xstring)
    id = RequiredAttribute("sheetId", XsdUnsignedInt)
    rel_id = RequiredAttribute("r:id", ST_RelationshipId)
    
    
