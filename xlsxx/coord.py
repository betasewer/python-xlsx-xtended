# encoding: utf-8

import re
import string


def column_to_index(ref):
    """
    カラムを示すアルファベットを0ベース序数に変換する。
    Params:
        column(str): A, B, C, ... Z, AA, AB, ...
    Returns:
        int: 0ベース座標
    """
    column = 0
    for i, ch in enumerate(reversed(ref)):
        d = string.ascii_uppercase.index(ch) + 1
        column += d * pow(len(string.ascii_uppercase),i)
    return column-1

def index_to_column(index):
    """
    0ベース序数をカラムを示すアルファベットに変換する。
    Params:
        index(int): 0ベース座標
    Returns:
        str: A, B, C, ... Z, AA, AB, ...
    """
    m = index + 1 # 1ベースにする
    k = 26
    digits = []
    while True:
        q = (m-1) // k
        d = m - q * k
        digit = string.ascii_uppercase[d-1]
        digits.append(digit)
        if m <= k:
            break
        m = q
    return "".join(reversed(digits))

#
#
# セル参照
#
#

# 正規表現
re_cellref = re.compile("([A-Z]+)([0-9]+)")

def ref_to_coord(ref):
    """
    セル参照を座標に変換する。
    Params:
        ref(str): セル参照
    Returns:
        Tuple: 行、列の0ベース座標 Tuple[int, int]
    """
    m = re_cellref.match(ref)
    if m is not None:
        column = column_to_index(m.group(1))
        row = int(m.group(2))-1
        return (row, column)
    return (None, None)

def coord_to_ref(coord):
    """
    座標をセル参照文字列に変換する。
    Params:
        coord(Tuple): 0ベース座標
    Returns:
        str: セル参照
    """
    row, column = coord
    c = index_to_column(column)
    r = str(row+1)
    return c + r

def split_ref(ref):
    """
    セル参照をセル文字と1ベース行番号文字に分割する。
    Params:
        ref(str):
    Returns:
        Tuple[str, str]: 列、行
    """
    m = re_cellref.match(ref)
    if m:
        return m.group(1), m.group(2)
    return None, None

def modify_ref(ref, row=None, col=None):
    """
    セル参照を修正して返す。
    Params:
        ref(str): もとのセル参照
        row(str/int): 1ベース行番号文字/0ベース行番号
        col(str/int): 列アルファベット/0ベース列番号
    """
    if ref:
        c, r = split_ref(ref)
    else:
        c, r = "", ""

    if row is not None:
        if isinstance(row, str):
            r = row
        elif isinstance(row, int):
            r = "{}".format(row + 1)

    if col is not None:
        if isinstance(col, str):
            c = col
        elif isinstance(col, int):
            c = index_to_column(col)

    return c + r

#
#
# セル範囲参照
#
#
def split_range_ref(ref):
    """
    セル参照範囲を開始と末尾参照に分割する。
    Params:
        ref(str):
    Returns:
        Tuple[str, str]:
    """
    beg, sep, end = ref.partition(":")
    if not sep:
        return None, None
    return beg, end

def range_ref_to_coord(ref):
    """
    範囲のセル参照から0ベース座標に変換する。
    Params:
        ref(str): セル範囲参照
    Returns:
        Tuple: 座標のタプル Tuple[Tuple[int, int], Tuple[int, int]]
    """
    c1, c2 = split_range_ref(ref)
    if c1 and c2:
        return ref_to_coord(c1), ref_to_coord(c2)
    else:
        return ref_to_coord(c1), None

def modify_range_ref(range, head=None, tail=None, headcol=None, tailcol=None, headrow=None, tailrow=None):
    """
    セル参照範囲を修正して返す。
    Params:
        range(str): もとのセル参照範囲
        head(str): 開始セル参照
        tail(str): 終了セル参照
        headcol(str): 開始列アルファベット
        tailcol(str): 終了列アルファベット
        headrow(str): 開始1ベース行番号
        tailrow(str): 終了1ベース行番号
    """
    h, t = split_range_ref(range)

    if head is not None:
        h = head
    if tail is not None:
        t = tail
    
    if headcol is not None and headrow is not None:
        h = modify_ref("", headcol, headrow)
    if tailcol is not None and tailrow is not None:
        t = modify_ref("", tailcol, tailrow)
    
    if headcol is not None:
        h = modify_ref(h, col=headcol)
    if tailcol is not None:
        t = modify_ref(t, col=tailcol)
    if headrow is not None:
        h = modify_ref(h, row=headrow)
    if tailrow is not None:
        t = modify_ref(t, row=tailrow)

    return "{}:{}".format(h, t)

#
#
# 整数座標組への変換インターフェース
#
#
def get_coord(coord):
    """
    セル参照、または座標タプルを0ベース座標に分解する。
    Params:
        coord(str/Tuple):
    Returns:
        Tuple:
    """
    if isinstance(coord, str):
        r, c = ref_to_coord(coord)
        return r, c
    else:
        return coord

def get_range_coord(range, coord2=None, *, rownum=None, columnnum=None):
    """
    セル範囲参照、または座標の組を0ベース座標に分解する。
    Params:
        range(str/Tuple): セル範囲参照／セル範囲座標の組／左上セル参照／左上セル座標
        coord2(str/Tuple): 右下セル参照／右下セル座標
        rownum(int): 左上からの行の増分
        columnnum(int): 左上からの列の増分
    Returns:
        Tuple:
    """
    if rownum or columnnum:
        # 左上セル ＋ 増分
        r1, c1 = get_coord(range)
        if r1 is None or c1 is None:
            raise TypeError("左上セルの指定に誤りがあります")
        r2 = r1 + (rownum or 1)
        c2 = c1 + (columnnum or 1)
        return (r1, c1), (r2, c2)

    if coord2 is None:
        # 範囲指定一つ
        if isinstance(range, str):
            c1, c2 = range_ref_to_coord(range)
            return c1, c2
        else:
            p1, p2 = range
            return p1, p2
        
    else:
        # 左上セルと右下セルの指定
        if not isinstance(coord2, (str, tuple)):
            raise TypeError("coord2")
        return get_coord(range), get_coord(coord2)



