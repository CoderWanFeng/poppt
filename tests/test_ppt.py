import unittest

import office

from poppt.api.ppt import *


class TestPPT(unittest.TestCase):
    def test_ppt2pdf(self):
        ppt2pdf(path=r'./ppt/', output_path=r'./ppt/test_pdf2')

    # def test_ppt2img(self):
    #     ppt2img(intput_path=r'D:\test\py310\ppt_test')
    def test_single_ppt2imgg(self):
        office.ppt.ppt2img(input_path=r'C:\Users\Lenovo\Desktop\temp\test\test\ppt\a',
                           output_path=r'./ppt/test_pdf2/imgs23',
                           img_type='jpg')
