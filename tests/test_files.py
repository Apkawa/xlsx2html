# coding: utf-8
from __future__ import unicode_literals

import os
from unittest import TestCase

FIXTURES_ROOT = os.path.join(os.path.dirname(__file__), 'fixtures')


class ArchiveFileFieldTestCase(TestCase):
    def setUp(self):
        self.xlsx_file = os.path.join(FIXTURES_ROOT, 'example.xslx')
