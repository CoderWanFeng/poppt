#!/usr/bin/env python
# -*- coding:utf-8 -*-

#############################################
# File Name: ppt.py
# 公众号/B站/小红书/抖音/知乎: 程序员晚枫
# Mail: 1957875073@qq.com
# Created Time:  2022-4-25 10:17:34
# Description: 有关 PPT 的自动化操作 API 接口层
#############################################
# 导入操作系统模块，用于路径转换
import os
# 导入 PPT 的核心处理类
from poppt.core.PPTType import MainPPT

# 创建 MainPPT 实例，用于处理所有 PPT 操作
mainPPT = MainPPT()


# todo：输入文件路径
# @except_dec()
def ppt2pdf(path: str, output_path='./'):
    """
    将 PPT/PPTX 文件转换为 PDF 格式

    Args:
        path: PPT 文件或文件夹路径
        output_path: PDF 文件的输出路径，默认为当前目录

    Returns:
        None
    """
    # 将路径转换为绝对路径
    path = os.path.abspath(path)
    output_path = os.path.abspath(output_path)
    mainPPT.ppt2pdf(path=path, output_path=output_path)


def ppt2img(input_path: str, output_path=r'./', merge: bool = False):
    """
    将 PPT 转换为图片，支持转为单独图片或合并为长图

    Args:
        input_path: 存放 PPT 的位置，
                    转换单个文件 → 可以写文件的路径
                    转换多个文件 → 写文件夹的路径
        output_path: 结果图片的存储位置，默认为代码目录
        merge: True → 将所有幻灯片合并为 1 张长图
               False → PPT 有多少张，就转为多少张独立图片

    Returns: 
        None

    """
    # 将路径转换为绝对路径
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    mainPPT.ppt2img(input_path=input_path,
                    output_path=output_path, merge=merge)


def merge4ppt(input_path, output_path=r'./', output_name='merge4ppt.pptx'):
    """
    合并多个 PPT 文件为一个完整的 PPT 文件

    Args:
        input_path: 包含多个 PPT 文件的文件夹路径
        output_path: 合并后 PPT 文件的输出路径，默认为当前目录
        output_name: 合并后 PPT 文件的名称，默认为 'merge4ppt.pptx'

    Returns:
        None
    """
    # 将路径转换为绝对路径
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    mainPPT.merge4ppt(input_path=input_path,
                      output_path=output_path, output_name=output_name)
