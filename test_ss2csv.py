#!/usr/bin/env python

#
# Tests
#

import imp
import sys
import pytest
from ss2csv import *

test_file = 'test_ss.xlsx'


def test_xlrd_installation():
    try:
        imp.find_module('xlrd')
        assert True
    except ImportError:
        sys.exit(1)


def test_is_spreadsheet():
    assert is_spreadsheet(test_file)


def test_workboot_data():
    wb = WorkbookData(test_file)
    assert isinstance(wb, WorkbookData)
    assert wb.worksheet_count >= 1


def test_save_to_csv():
    wb = WorkbookData(test_file)
    s = SaveToCSV(wb)
    assert isinstance(s, threading.Thread)
    assert isinstance(s, SaveToCSV)
