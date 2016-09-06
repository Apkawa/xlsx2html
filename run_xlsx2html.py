#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys

from xlsx2html import xls2html

if __name__ == '__main__':
    xls2html(sys.argv[1], sys.argv[2])


