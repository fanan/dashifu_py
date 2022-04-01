#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import os.path
import shutil
import xlwings
import sys

filelists = []


def is_excel(fn):
    _, ext = os.path.splitext(fn)
    if ext == ".xls" or ext == ".xlsx":
        if "~$" not in fn and "新建 Microsoft Excel 工作表" not in fn:
            return True
    return False


def refresh(fn):
    tmp_xls = os.path.expanduser("~/tmp/1.xls")
    tmp_xlsx = os.path.expanduser("~/tmp/1.xlsx")
    if fn.endswith(".xlsx"):
        tmp_fn = tmp_xlsx
    else:
        tmp_fn = tmp_xls
    shutil.copy(fn, tmp_fn)
    wb = xlwings.Book(tmp_fn)
    wb.save()
    wb.close()
    os.rename(tmp_fn, fn)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        root_dir = sys.argv[1]
    else:
        root_dir = os.path.expanduser("~/Downloads/dashifu")

    for root, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if is_excel(filename):
                filelists.append(os.path.join(root, filename))

    for fn in filelists:
        print("refreshing {}".format(fn))
        refresh(fn)
