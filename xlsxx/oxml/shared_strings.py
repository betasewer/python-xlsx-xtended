# encoding: utf-8

"""
"""

from docxx.oxml.xmlchemy import (
    BaseOxmlElement, ZeroOrOne, ZeroOrMore, OneOrMore, OneAndOnlyOne,
    RequiredAttribute, OptionalAttribute, 
)
from docxx.oxml.simpletypes import (
    XsdUnsignedInt, ST_RelationshipId
)
from xlsxx.oxml.simpletypes import (
    ST_Xstring, ST_FontId, ST_PhoneticAlignment, ST_PhoneticType
)


#
#
#
class CT_SST(BaseOxmlElement):
    si = ZeroOrMore('ssml:si')
    extLst = ZeroOrOne('ssml:extLst') # None
    
    count = OptionalAttribute('count', XsdUnsignedInt) # String Count
    uniqueCount = OptionalAttribute('uniqueCount', XsdUnsignedInt) # Unique String Count

class CT_RST(BaseOxmlElement):
    t = ZeroOrOne('ssml:t')
    r = ZeroOrMore('ssml:r')
    
    rPh = ZeroOrMore('ssml:rPh')
    phoneticPr = ZeroOrOne('ssml:phoneticPr')

class CT_PhoneticRun(BaseOxmlElement):
    """
    Complex type: 
    """
    t = OneAndOnlyOne('ssml:t') # Text
    
    sb = RequiredAttribute('sb', XsdUnsignedInt) # Base Text Start Index
    eb = RequiredAttribute('eb', XsdUnsignedInt) # Base Text End Index
    
class CT_PhoneticPr(BaseOxmlElement):
    """
    Complex type: 
    """
    fontId = RequiredAttribute('fontId', ST_FontId) # Font Id
    type = OptionalAttribute('type', ST_PhoneticType) # Character Type
    alignment = OptionalAttribute('alignment', ST_PhoneticAlignment) # Alignment
    
class CT_RElt(BaseOxmlElement):
    """
    Complex type: 
    """
    rPr = ZeroOrOne('ssml:rPr') # Run Properties
    t = OneAndOnlyOne('ssml:t') # Text
    
    
class CT_RPrElt(BaseOxmlElement):
    """
    Complex type: 
    """
    rFont = ZeroOrOne('ssml:rFont') # Font
    charset = ZeroOrOne('ssml:charset') # Character Set
    family = ZeroOrOne('ssml:family') # Font Family
    b = ZeroOrOne('ssml:b') # Bold
    i = ZeroOrOne('ssml:i') # Italic
    strike = ZeroOrOne('ssml:strike') # Strike Through
    outline = ZeroOrOne('ssml:outline') # Outline
    shadow = ZeroOrOne('ssml:shadow') # Shadow
    condense = ZeroOrOne('ssml:condense') # Condense
    extend = ZeroOrOne('ssml:extend') # Extend
    color = ZeroOrOne('ssml:color') # Text Color
    sz = ZeroOrOne('ssml:sz') # Font Size
    u = ZeroOrOne('ssml:u') # Underline
    vertAlign = ZeroOrOne('ssml:vertAlign') # Vertical Alignment
    scheme = ZeroOrOne('ssml:scheme') # Font Scheme
    
