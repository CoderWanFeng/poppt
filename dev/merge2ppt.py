# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/yFcocJbfS9Hs375NhE8Gbw
@个人网站 ：www.python-office.com
@Date    ：2023/5/25 22:44 
@Thanks     ：https://zhuanlan.zhihu.com/p/434485047
'''
from pathlib import Path

import win32com.client as win32
from pofile import mkdir, get_files
from poprogress import simple_progress


def merge4ppt(input_path: str, output_path: str, output_name: str):
    """
    :param path: ppt所在文件路径
    :return: None
    """

    abs_input_path = Path(input_path).absolute()  # 相对路径→绝对路径
    exsit, abs_output_path = mkdir(output_path)
    ppt_file_list = get_files(abs_input_path)
    save_path = Path(abs_output_path) / output_name

    Application = win32.gencache.EnsureDispatch("PowerPoint.Application")

    Application.Visible = 1
    new_ppt = Application.Presentations.Add()
    # 执行合并操作
    for ppt_file in simple_progress(ppt_file_list):
        exit_ppt = Application.Presentations.Open(ppt_file)
        print('正在操作的文件：', ppt_file)
        page_num = exit_ppt.Slides.Count
        exit_ppt.Close()
        new_ppt.Slides.InsertFromFile(ppt_file, new_ppt.Slides.Count, 1, page_num)
    new_ppt.Save()  # 括号内为保存位置：如C:\Users\Administrator\Documents\下
    Application.Quit()


# join_ppt(r"E:\Downloads\PPT")

input_path = r"D:\workplace\code\github\poppt\dev\docs"
output_path = r"D:\workplace\code\github\poppt\dev\docs\output"
output_name = "test.pptx"

merge4ppt(input_path, output_path, output_name)
