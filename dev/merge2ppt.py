# -*- coding: UTF-8 -*-
'''
@作者：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@读者群：http://www.python4office.cn/wechat-group/
@个人网站：www.python-office.com
@Date：2023/5/25 22:44 
@Thanks：https://zhuanlan.zhihu.com/p/434485047
@Description：合并 PPT 的开发测试代码
'''
# 导入路径操作模块
from pathlib import Path

# 导入 Windows COM 组件
import win32com.client as win32
# 导入文件操作工具函数
from pofile import mkdir, get_files
# 导入进度条工具
from poprogress import simple_progress


def merge4ppt(input_path: str, output_path: str, output_name: str):
    """
    合并多个 PPT 文件为一个完整的 PPT 文件
    
    Args:
        input_path: PPT 文件所在的文件夹路径
        output_path: 合并后 PPT 文件的输出路径
        output_name: 合并后 PPT 文件的名称
    
    Returns:
        None
    """

    # 将相对路径转换为绝对路径
    abs_input_path = Path(input_path).absolute()  # 相对路径→绝对路径
    # 创建输出目录，返回是否存在和绝对路径
    exsit, abs_output_path = mkdir(output_path)
    # 获取文件夹下所有 PPT 文件列表
    ppt_file_list = get_files(abs_input_path)
    # 生成完整的保存路径
    save_path = Path(abs_output_path) / output_name

    # 创建 PowerPoint 应用程序实例
    Application = win32.gencache.EnsureDispatch("PowerPoint.Application")

    # 设置 PowerPoint 可见，1 表示显示，0 表示隐藏
    Application.Visible = 1
    # 创建一个新的 PPT 文件，用于存放合并后的内容
    new_ppt = Application.Presentations.Add()
    # 执行合并操作
    for ppt_file in simple_progress(ppt_file_list):
        # 打开当前 PPT 文件
        exit_ppt = Application.Presentations.Open(ppt_file)
        print('正在操作的文件：', ppt_file)
        # 获取当前 PPT 的幻灯片数量
        page_num = exit_ppt.Slides.Count
        # 关闭当前 PPT 文件
        exit_ppt.Close()
        # 将当前 PPT 的所有幻灯片插入到新 PPT 的末尾
        new_ppt.Slides.InsertFromFile(ppt_file, new_ppt.Slides.Count, 1, page_num)
    # 保存合并后的 PPT 文件（括号内为保存位置：如 C:\Users\Administrator\Documents\下）
    new_ppt.Save()  # 括号内为保存位置：如C:\Users\Administrator\Documents\下
    # 退出 PowerPoint 应用程序
    Application.Quit()


# join_ppt(r"E:\Downloads\PPT")

# 设置要合并的 PPT 文件所在的文件夹路径
input_path = r"D:\workplace\code\github\poppt\dev\docs"
# 设置合并后 PPT 文件的输出路径
output_path = r"D:\workplace\code\github\poppt\dev\docs\output"
# 设置合并后 PPT 文件的名称
output_name = "test.pptx"

# 调用 merge4ppt 函数执行合并操作
merge4ppt(input_path, output_path, output_name)
