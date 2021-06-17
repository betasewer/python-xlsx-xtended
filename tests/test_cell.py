import pytest
from xlsxx.proxy.cell import (
    Cell, CellRange, CellRow, 
    WINDOWS_EXCEL_TIME, OLD_MAC_EXCEL_TIME, OTHER_APP_TIME
)
import datetime

def test_excel_to_datetime():
    # windows
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(40729) == datetime.datetime(2011, 7, 5)
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(12345) == datetime.datetime(1933, 10, 18)
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(1234) == datetime.datetime(1903, 5, 18)
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(123) == datetime.datetime(1900, 5, 2)
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(61) == datetime.datetime(1900, 3, 1)
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(60) == datetime.datetime(1900, 2, 28)
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(59) == datetime.datetime(1900, 2, 28)
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(1) == datetime.datetime(1900, 1, 1)
    assert WINDOWS_EXCEL_TIME.convert_to_datetime(0) == datetime.datetime(1899, 12, 30)
    # other (libreoffice)
    assert OTHER_APP_TIME.convert_to_datetime(40729) == datetime.datetime(2011, 7, 5)
    assert OTHER_APP_TIME.convert_to_datetime(61) == datetime.datetime(1900, 3, 1)
    assert OTHER_APP_TIME.convert_to_datetime(60) == datetime.datetime(1900, 2, 28)
    assert OTHER_APP_TIME.convert_to_datetime(59) == datetime.datetime(1900, 2, 27)
    assert OTHER_APP_TIME.convert_to_datetime(1) == datetime.datetime(1899, 12, 31)
    assert OTHER_APP_TIME.convert_to_datetime(0) == datetime.datetime(1899, 12, 30)

def test_excel_to_time():
    # アプリごとの差はない
    assert WINDOWS_EXCEL_TIME.convert_to_time(0.440277777777778) == datetime.time(10, 34, 00)
    assert WINDOWS_EXCEL_TIME.convert_to_time(0.982465277777778) == datetime.time(23, 34, 45)
    assert WINDOWS_EXCEL_TIME.convert_to_time(0) == datetime.time(0, 0, 0)
