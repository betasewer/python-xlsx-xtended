import os
from typing import Callable

from docxx.api import DocumentOpener
from docxx.opc.constants import CONTENT_TYPE as CT

from xlsxx.parts.sml import SmlSheetMainPart

def _templatefile(filename):
    thisdir = os.path.split(__file__)[0]
    return os.path.join(thisdir, 'templates', filename)

# ファイルオープン関数
open_xlsx: DocumentOpener[SmlSheetMainPart] = DocumentOpener(
    CT.SML_SHEET_MAIN, "Excel File", _templatefile("default.xlsx")
)

# シートを取得する
def open_xlsx_sheet(path, sheet=None):
    xlsx = open_xlsx(path)
    if sheet is None:
        sht = xlsx.workbook.topsheet 
    else:
        sht = xlsx.workbook.sheet(index=sheet)
    return sht

def read_xlsx_vertical(path, start, tailcolumn, tailrow=-1, *, sheet=None, sequence=False, nofetch=False):
    """
    シートから縦方向に値を読みだす
    Params:
        path(Path): 対象ファイル
        start(Tuple[str|int, str|int]|str): 開始位置の座標
        tailcolumn(str|int): 終了位置の列参照・列番号（含む）
        tailrow(str|int) : 終了位置の列参照・列番号（含む）
        sheet(int): シートのインデックス
        sequence(bool): 空の行があったら終了する
        nofetch(bool): 一度にテキストを取得しない
    Returns:
        List[List[Str]]: 文字列型の値が列の数だけ入った行のリスト
    """
    from xlsxx.proxy.sheet import read_sheet_vertical
    sht = open_xlsx_sheet(path, sheet)
    return read_sheet_vertical(sht, start, tailcolumn, tailrow, sequence=sequence, nofetch=nofetch)
