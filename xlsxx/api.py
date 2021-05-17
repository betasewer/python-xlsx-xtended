import os

from docxx.api import DocumentOpener
from docxx.opc.constants import CONTENT_TYPE as CT

def _templatefile(filename):
    thisdir = os.path.split(__file__)[0]
    return os.path.join(thisdir, 'templates', filename)

# ファイルオープン関数
open_xlsx = DocumentOpener(CT.SML_SHEET_MAIN, "Excel File", _templatefile("default.xlsx"))

