from pathlib import *

import win32com.client
from pofile import get_files, mkdir
from poprogress import simple_progress
from poppt.lib.ppt.ppt2pdf_service import ppt2pdf_single


class MainPPT():

    def ppt2pdf(self, path, output_path):
        """
        @Author & Date  : CoderWanFeng 2022/5/9 23:34
        @Desc  : path:存放ppt的路径
        """
        filenames = get_files(path)
        exsit, abs_output_path = mkdir(output_path)
        for filename in simple_progress(filenames):
            # 判断文件的类型，对所有的ppt文件进行处理(ppt文件以ppt或者pptx结尾的)
            if filename.endswith('ppt') or filename.endswith('pptx'):
                new_pdf_name = Path(filename).stem + '.pdf'  # PPT素材1.pdf
                output_pdf_filename = Path(abs_output_path).absolute() / new_pdf_name
                ppt2pdf_single(filename, str(output_pdf_filename))
                # time.sleep(3)

    def ppt2img(self, input_path, output_path, img_type):
        '''将PPT另存为图片格式
          arguments:
              pptFullName: 要转换的ppt文件，
              pptName：转换后的存放JPG文件的目录
              imgType: 图片类型
        '''
        filenames = get_files(input_path)
        # 启动PPT
        pptClient = win32com.client.Dispatch('PowerPoint.Application')
        # 设置为0表示后台运行，不显示，1则显示
        pptClient.Visible = True
        for ppt_file in filenames:
            if ppt_file.endswith('ppt') or ppt_file.endswith('pptx'):
                # Python路径操作模块pathlib，看这篇就够了！ https://zhuanlan.zhihu.com/p/475661402
                output_dir = Path(output_path).absolute() / str(Path(ppt_file).stem)
                exsit, output_dir = mkdir(output_dir)
                # JPG是17
                img_type_jpg = 17
                # PNG是18
                img_type_png = 18

                # 打开PPT文件
                ppt = pptClient.Presentations.Open(ppt_file, WithWindow=False, ReadOnly=1)
                # 另存为图片
                if img_type == 'jpg':
                    ppt.SaveAs(output_dir, img_type_jpg)
                else:
                    ppt.SaveAs(output_dir, img_type_png)
        # 退出
        pptClient.Quit()

    def merge4ppt(self, input_path: str, output_path: str, output_name: str):
        """
        :param path: ppt所在文件路径
        :return: None
        """

        abs_input_path = Path(input_path).absolute()  # 相对路径→绝对路径
        exsit, abs_output_path = mkdir(output_path)
        ppt_file_list = get_files(abs_input_path)
        save_path = Path(abs_output_path) / output_name

        Application = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")

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
