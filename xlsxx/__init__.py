# encoding: utf-8
""" @module
Extra-Requirements:
    docxx
DefModules:
    xlsxx.ma
"""

__version__ = '0.0.1.0'

from docxx.opc.constants import CONTENT_TYPE as CT
from docxx.opc.part import PartFactory

from xlsxx.parts.sml import SmlSheetMainPart, SmlWorkSheetPart, SmlSharedStringsPart, SmlStylesPart

# xlsx
PartFactory.part_type_for[CT.SML_SHEET_MAIN] = SmlSheetMainPart
PartFactory.part_type_for[CT.SML_WORKSHEET] = SmlWorkSheetPart
PartFactory.part_type_for[CT.SML_SHARED_STRINGS] = SmlSharedStringsPart
PartFactory.part_type_for[CT.SML_STYLES] = SmlStylesPart

del (
    SmlSheetMainPart, 
    SmlWorkSheetPart, 
    SmlSharedStringsPart, 
    SmlStylesPart
)
