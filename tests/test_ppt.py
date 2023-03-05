import unittest

from poppt.api.ppt import *


class TestPPT(unittest.TestCase):
    def test_ppt2pdf(self):
        ppt2pdf(path=r'./ppt/')

    # def test_ppt2img(self):
    #     ppt2img(intput_path=r'D:\test\py310\ppt_test')
    def test_single_ppt2imgg(self):
        ppt2img(intput_path=r'D:\test\py310\ppt_test\代码之外，程序员能力提升必备的8个软件（刘兆锋）.pptx',
                output_path=r'D:\test\py310\ppt_test',
                img_type='jpg')
