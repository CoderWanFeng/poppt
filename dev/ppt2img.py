# -*- coding: UTF-8 -*-
'''
@作者：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@读者群：http://www.python4office.cn/wechat-group/
@个人网站：www.python-office.com
@Date：2023/6/8 23:33 
@Description：PPT 转图片的开发测试代码
'''

# -*- coding: UTF-8 -*-
# Auther: youren.S

# 导入系统操作相关模块
import os
import shutil

# 导入 Windows COM 组件
import win32com
import win32com.client
# from win32com.client import constants
# 导入图片处理库
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
    将 PPT 转换为 PNG 图片格式
    
    Args:
        ppt_path: PPT 文件的绝对路径
        long_sign: 是否需要转为生成长图的标识（当前未使用）
    
    Returns:
        None
    
    Raises:
        Exception: 如果文件不存在，抛出异常
    """
    # 检查 PPT 文件是否存在
    if os.path.exists(ppt_path):
        # 判断文件是否存在并生成图片保存目录
        output_path = output_file(ppt_dirfile)  # 判断文件是否存在
        # 创建 PowerPoint 应用程序实例
        ppt_app = win32com.client.Dispatch('PowerPoint.Application')
        # 设置为 0 表示后台运行，不显示，1 则显示
        ppt_app.Visible = 1
        # 打开 PPT 文件
        ppt = ppt_app.Presentations.Open(ppt_path)  # 打开 ppt
        # 17 表示 PPT 转为 JPG 图片
        ppSaveAsJPG = 17
        # 将 PPT 另存为图片格式
        ppt.SaveAs(output_path, ppSaveAsJPG)  # 17表示 ppt 转为图片
        # 关闭资源，退出 PowerPoint 应用程序
        ppt_app.Quit()  # 关闭资源，退出
        # 合并生成长图
        generate_long_image(output_path)  # 合并生成长图

    else:
        # 如果文件不存在，抛出异常
        raise Exception('请检查文件是否存在！\n')


def generate_long_image(output_path):
    """
    将 PPT 的各个页面拼接成长图
    
    Args:
        output_path: PPT 页面图片存放的路径
    
    Returns:
        None
    """
    # 获取图片目录路径
    picture_path = output_path
    # 获取上一级文件目录
    last_dir = os.path.dirname(picture_path)  # 上一级文件目录

    # 获取图片列表
    # ims = [Image.open(os.path.join(picture_path, fn)) for fn in os.listdir(picture_path) if fn.endswith('.jpg')]
    ims = []
    # 遍历目录下所有文件
    for fn in os.listdir(picture_path):
        # 只处理 JPG 格式的图片文件（忽略大小写）
        if fn.lower().endswith('.jpg'):
            ims.append(os.path.join(picture_path, fn))

    # print(ims)
    # 将获取到 PPT 的页面进行排序，按照文件名长度排序
    ims_sort = sorted(ims, key=lambda jpg: len(jpg))
    print(ims_sort)

    # 取第一个图片的尺寸（宽度和高度）
    width, height = Image.open(ims[0]).size  # 取第一个图片尺寸
    # 创建同宽、n 倍高的空白图片，用于存放长图
    long_canvas = Image.new(Image.open(ims[0]).mode, (width, height * len(ims)))  # 创建同宽，n倍高的空白图片

    # 拼接图片，将每个页面图片依次粘贴到长画布上
    for i, image in enumerate(ims_sort):
        # 将图片粘贴到长画布的指定位置，从上到下排列
        long_canvas.paste(Image.open(image), box=(0, i * height))
    # 保存生成的长图为 PNG 格式
    long_canvas.save(os.path.join(last_dir, ppt_name1 + '.png'))  # 保存长图


if __name__ == '__main__':
    # # ppt_path = "test_template.pptx"
    # cur_path = os.getcwd()
    # print(cur_path)
    # ppt_path = os.path.join(cur_path, ppt_dirfile)  # 需要使用绝对路径，否则会报错
    # long_sign = "y"
    # 设置要转换的 PPT 文件路径
    ppt_path = r'D:\workplace\code\github\poppt\tests\ppt\test_ppt_img'
    # 设置是否需要合并为长图
    long_sign = True
    # 调用 ppt2png 函数执行转换
    ppt2png(ppt_path, long_sign)
