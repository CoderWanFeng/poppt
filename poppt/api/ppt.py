#!/usr/bin/env python
# -*- coding:utf-8 -*-

#############################################
# File Name: ppt.py
# 公众号/B站/小红书/抖音/知乎: 程序员晚枫
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


def ppt2img(input_path: str, output_path=r'./', merge: bool = False):
    """
    PPT转图片，可以转为长图
    Args:
        input_path: 存放PPT的位置，
                    转换单个文件 → 可以写文件的路径
                    转换单个文件 → 写文件夹的路径
        output_path: 结果图片的存储位置，可以不写，默认代码目录
        merge: True → 转为1张图片
            False → PPT有多少张，就转为多少张图片

    Returns: None

    """
    mainPPT.ppt2img(input_path, output_path, merge)


def merge4ppt(input_path, output_path=r'./', output_name='merge4ppt.pptx'):
    mainPPT.merge4ppt(input_path, output_path, output_name)
