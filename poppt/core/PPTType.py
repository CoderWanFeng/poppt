import os
import shutil
from pathlib import *

import win32com.client
from PIL import Image
from pofile import get_files, mkdir, check_suffix
from poprogress import simple_progress

from poppt.lib.ppt.ppt2pdf_service import ppt2pdf_single


class MainPPT():
    def __init__(self):
        self.app = 'PowerPoint.Application'
        self.suffix_list = ["ppt", "pptx"]
        self.default_img_type = ".jpg"
        # JPG是17，PNG是18
        self.ppt2img_type = 17

    def ppt2pdf(self, path, output_path):
        """
        @Author & Date  : CoderWanFeng 2022/5/9 23:34
        @Desc  : path:存放ppt的路径
        """
        filenames = get_files(path)
        exsit, abs_output_path = mkdir(output_path)
        for filename in simple_progress(filenames):
            # 判断文件的类型，对所有的ppt文件进行处理(ppt文件以ppt或者pptx结尾的)
            if check_suffix(filename, self.suffix_list):
                new_pdf_name = Path(filename).stem + '.pdf'  # PPT素材1.pdf
                output_pdf_filename = Path(abs_output_path).absolute() / new_pdf_name
                ppt2pdf_single(filename, str(output_pdf_filename))
                # time.sleep(3)

    def merge4ppt(self, input_path: str, output_path: str, output_name: str):
        """
        :param path: ppt所在文件路径
        :return: None
        """
        abs_input_path = Path(input_path).absolute()  # 相对路径→绝对路径
        exsit, abs_output_path = mkdir(output_path)
        ppt_file_list = get_files(abs_input_path)
        save_path = Path(abs_output_path) / output_name

        Application = win32com.client.gencache.EnsureDispatch(self.app)

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

    def ppt2img(self, input_path, output_path, merge=False):
        '''将PPT另存为图片格式
          arguments:
              pptFullName: 要转换的ppt文件，
              pptName：转换后的存放JPG文件的目录
        '''

        filenames = get_files(input_path)
        # 启动PPT
        pptClient = win32com.client.Dispatch(self.app)
        # 设置为0表示后台运行，不显示，1则显示
        pptClient.Visible = True
        for ppt_file in filenames:
            if check_suffix(ppt_file, self.suffix_list):
                # Python路径操作模块pathlib，看这篇就够了！ https://zhuanlan.zhihu.com/p/475661402
                output_dir = Path(output_path).absolute() / str(Path(ppt_file).stem)
                exsit, output_dir = mkdir(output_dir)
                # 打开PPT文件
                ppt = pptClient.Presentations.Open(ppt_file, WithWindow=False, ReadOnly=1)
                # 另存为图片
                ppt.SaveAs(output_dir, self.ppt2img_type)

        # 退出
        pptClient.Quit()
        if merge:
            merge_dir = Path(output_path).absolute()
            for dirpath, dirnames, filenames in os.walk(merge_dir):
                # print(dirpath, dirnames, filenames)
                for dirname in dirnames:
                    current_img_path = os.path.join(dirpath, dirname)
                    self.generate_long_image(input_path=current_img_path,
                                             current_ppt_name=dirname,
                                             output_path=merge_dir,
                                             img_name=dirname + self.default_img_type)
                    shutil.rmtree(current_img_path)  # 删除暂存图片

    def generate_long_image(self, input_path: str, current_ppt_name, output_path, img_name='merge.jpg'):
        """
        将ppt的各个页面拼接成长图：https://blog.csdn.net/m0_51777056/article/details/130262561
        :param input_path:
        :param output_path:
        :param img_name:
        :return:
        """
        # 获取图片列表
        img_list = []
        for imgs in os.listdir(input_path):
            img_list.append(os.path.join(input_path, imgs))

        # 将获取到ppt的页面进行排序
        ims_sort = sorted(img_list, key=lambda jpg: len(jpg))
        print(f"正在生成【PPT文件名为：{current_ppt_name}】的图片")

        width, height = Image.open(img_list[0]).size  # 取第一个图片尺寸
        img_mode = Image.open(img_list[0]).mode
        long_canvas = Image.new(img_mode, (width, height * len(img_list)))  # 创建同宽，n倍高的空白图片

        # 拼接图片
        for i, image in enumerate(ims_sort):
            long_canvas.paste(Image.open(image), box=(0, i * height))
        long_canvas.save(os.path.join(output_path, img_name))  # 保存长图
