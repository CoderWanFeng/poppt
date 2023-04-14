#!/usr/bin/env python
# -*- coding:utf-8 -*-

#############################################
# File Name: ppt.py
# Mail: 1957875073@qq.com
# Created Time:  2022-4-25 10:17:34
# Description: 有关 ppt 的自动化操作
#############################################
from poppt.core.PPTType import MainPPT

mainPPT = MainPPT()


# todo：输入文件路径
# @except_dec()
def ppt2pdf(path: str, output_path='./'):
    mainPPT.ppt2pdf(path, output_path)


# def ppt2img(intput_path, output_path=r'./', img_type='img'):
#     mainPPT.ppt2img(intput_path, output_path, img_type)
def ppt2img(input_path, output_path, img_type):
    """
    :param intput_path:
    :param output_path:
    :param img_type:
    :return:
    """
    mainPPT.ppt2img(input_path, output_path, img_type)
