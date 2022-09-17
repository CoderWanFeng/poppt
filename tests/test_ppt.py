import unittest

from poppt.api.ppt import *


class TestPPT(unittest.TestCase):
    def test_ppt2pdf(self):
        ppt2pdf(path=r'./ppt/')
