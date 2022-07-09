from json import detect_encoding
from xlsxx.proxy.styles import (
    NUMVAL_TYPE_DATETIME,
    NUMVAL_TYPE_TIME,
    NUMVAL_TYPE_FLOAT,
    NUMVAL_TYPE_INT,
    numfmt_lib,
)

def test_numformat_detection():
    detect = numfmt_lib.detect_value_type
    convert = numfmt_lib.make_valtype_detection_string

    assert detect("Standard") == NUMVAL_TYPE_INT
    
    # 日付
    assert detect("YYYY/M/D;@") == NUMVAL_TYPE_DATETIME
    assert detect("YY年 QQ") == NUMVAL_TYPE_DATETIME

    # 時刻
    assert detect("H:MM:SS AM/PM") == NUMVAL_TYPE_TIME
    assert detect("[HH]:MM:SS.00") == NUMVAL_TYPE_TIME
    
    # パーセント
    assert detect("0.00%") == NUMVAL_TYPE_FLOAT

    # 数値
    assert detect("#,##0") == NUMVAL_TYPE_INT
    assert convert("#,##0;[RED]-#,##0") == "#,##0; -#,##0"
    assert detect("#,##0;[RED]-#,##0") == NUMVAL_TYPE_INT
    assert detect("¥#,##0.00;[RED]-¥#,##0.00") == NUMVAL_TYPE_FLOAT
    assert detect("#,##0.00;[RED]-#,##0.00") == NUMVAL_TYPE_FLOAT
    assert convert("¥#,##0.00;[RED]-¥#,##0.00") == "¥#,##0.00; -¥#,##0.00"
    assert detect("$#,##0.00_);($#,##0.00)") == NUMVAL_TYPE_FLOAT

    # 通貨
    assert detect("[$￥-411]#,##0.00;[RED]-[$￥-411]#,##0.00") == NUMVAL_TYPE_FLOAT
    assert detect("[$￥-411]#,##0.--;[RED]-[$￥-411]#,##0.--") == NUMVAL_TYPE_FLOAT
    assert convert("[$￥-411]#,##0.00;[RED]-[$￥-411]#,##0.00") == " #,##0.00; - #,##0.00"

    # 指数表記
    assert detect("0.00E+00") == NUMVAL_TYPE_FLOAT

    # 分数
    assert detect("# ???/???") == NUMVAL_TYPE_FLOAT

