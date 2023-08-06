# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/yFcocJbfS9Hs375NhE8Gbw
@个人网站 ：www.python-office.com
@Date    ：2023/6/8 23:33 
@Description     ：
'''

# -*- coding: UTF-8 -*-
# Auther: youren.S

import os
import shutil

import win32com
import win32com.client
# from win32com.client import constants
from PIL import Image
#
# ppt_dir = r'D:\workplace\code\github\poppt\tests\ppt\test_ppt_img'
# ppt_file = 'test_pdf2.pptx'
# ppt_name1, ppt_name2 = ppt_file.split(".")
#
# file_ext = os.path.splitext(ppt_file)[1]
#
# # 判断文件后缀
# if file_ext == ".ppt" or file_ext == ".pptx":
#     print("=" * 60)
# else:
#     print("请检查文件后缀是否正确。")
#
# ppt_dirfile = ppt_dir + '/' + ppt_file
# print(ppt_dirfile)


# wc = win32com.client.constants
# def output_file(ppt_path):
#     """
#     判断文件是否存在并生成图片保存目录
#     :param ppt_path: ppt文件路径
#     :return: 文件保存目录
#     """
#     output_path = ""
#     if os.path.exists(ppt_path):
#         fname, ext = os.path.splitext(ppt_file)
#         output_path = os.path.join(ppt_dir, fname)
#         if os.path.isdir(output_path):
#             shutil.rmtree(output_path)
#         os.mkdir(output_path)
#     return output_path


def ppt2png(ppt_path, ):
    """
    ppt 转 png 方法
    :param ppt_path: ppt 文件的绝对路径
    :param long_sign: 是否需要转为生成长图的标识
    :return:
    """
    if os.path.exists(ppt_path):
        output_path = output_file(ppt_dirfile)  # 判断文件是否存在
        ppt_app = win32com.client.Dispatch('PowerPoint.Application')
        # 设置为0表示后台运行，不显示，1则显示
        ppt_app.Visible = 1
        ppt = ppt_app.Presentations.Open(ppt_path)  # 打开 ppt
        ppSaveAsJPG = 17
        ppt.SaveAs(output_path, ppSaveAsJPG)  # 17表示 ppt 转为图片
        ppt_app.Quit()  # 关闭资源，退出
        generate_long_image(output_path)  # 合并生成长图

    else:
        raise Exception('请检查文件是否存在！\n')


def generate_long_image(output_path):
    """
    将ppt的各个页面拼接成长图
    :param output_path:
    :return:
    """
    picture_path = output_path
    last_dir = os.path.dirname(picture_path)  # 上一级文件目录

    # 获取图片列表
    # ims = [Image.open(os.path.join(picture_path, fn)) for fn in os.listdir(picture_path) if fn.endswith('.jpg')]
    ims = []
    for fn in os.listdir(picture_path):
        if fn.lower().endswith('.jpg'):
            ims.append(os.path.join(picture_path, fn))

    # print(ims)
    # 将获取到ppt的页面进行排序
    ims_sort = sorted(ims, key=lambda jpg: len(jpg))
    print(ims_sort)

    width, height = Image.open(ims[0]).size  # 取第一个图片尺寸
    long_canvas = Image.new(Image.open(ims[0]).mode, (width, height * len(ims)))  # 创建同宽，n倍高的空白图片

    # 拼接图片
    for i, image in enumerate(ims_sort):
        long_canvas.paste(Image.open(image), box=(0, i * height))
    long_canvas.save(os.path.join(last_dir, ppt_name1 + '.png'))  # 保存长图


if __name__ == '__main__':
    # # ppt_path = "test_template.pptx"
    # cur_path = os.getcwd()
    # print(cur_path)
    # ppt_path = os.path.join(cur_path, ppt_dirfile)  # 需要使用绝对路径，否则会报错
    # long_sign = "y"
    ppt_path = r'D:\workplace\code\github\poppt\tests\ppt\test_ppt_img'
    long_sign = True
    ppt2png(ppt_path, long_sign)
