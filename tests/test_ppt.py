import unittest

import office

from poppt.api.ppt import *


class TestPPT(unittest.TestCase):
    def test_ppt2pdf(self):
        ppt2pdf(path=r'./ppt/test_ppt_img', output_path=r'./ppt/test_ppt_img/imgs')

    # def test_ppt2img(self):
    #     ppt2img(intput_path=r'D:\test\py310\ppt_test')
    def test_single_ppt2imgg(self):
        office.ppt.ppt2img(input_path=r'C:\Users\Lenovo\Desktop\temp\test\test\ppt\a',
                           output_path=r'./ppt/test_pdf2/imgs23',
                           img_type='jpg')
    def test_mergeppt(self):
        office.ppt.merge4ppt(input_path=r'd:\\程序员晚枫的文件夹', output_path=r'./', output_name='merge4ppt.pptx')
