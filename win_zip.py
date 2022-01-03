#!/usr/bin/env python3
# coding=utf-8

from __future__ import absolute_import
from __future__ import print_function
import sys
import os
import zipfile
import six

def convert_utf8_to_gb18030(fn):
    return six.ensure_text(fn).encode("gb18030")

def zipdir(path, zf):
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(".xlsx"):
                fn = os.path.join(root, file)
                zf.write(fn, convert_utf8_to_gb18030(fn))

output = sys.argv[1]
with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
    zipdir(".", zf)
