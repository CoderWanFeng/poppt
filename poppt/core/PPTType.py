# 导入系统操作相关模块
import os
import shutil
# 导入路径操作模块
from pathlib import *

# 导入 Windows COM 组件，用于操作 PowerPoint
import win32com.client
# 导入图片处理库
from PIL import Image
# 导入文件操作工具函数
from pofile import get_files, mkdir, check_suffix
# 导入进度条工具
from poprogress import simple_progress

# 导入 PPT 转 PDF 的服务函数
from poppt.lib.ppt.ppt2pdf_service import ppt2pdf_single


class MainPPT():
    """
    PPT 操作的主类，封装了 PPT 转 PDF、PPT 转图片、合并 PPT 等功能
    """
    def __init__(self):
        """
        初始化 MainPPT 类，设置默认配置参数
        """
        # PowerPoint 应用程序的 COM 组件名称
        self.app = 'PowerPoint.Application'
        # 支持的 PPT 文件后缀名列表
        self.suffix_list = ["ppt", "pptx"]
        # 默认输出图片的格式
        self.default_img_type = ".jpg"
        # JPG 是 17，PNG 是 18
        # PPT 转图片的类型，17 表示 JPG 格式
        self.ppt2img_type = 17

    def ppt2pdf(self, path, output_path):
        """
        将 PPT 文件批量转换为 PDF 格式
        
        Args:
            path: 存放 PPT 的路径，可以是单个文件或文件夹
            output_path: 转换后 PDF 文件的输出路径
        
        Returns:
            None
            
        Author & Date: CoderWanFeng 2022/5/9 23:34
        """
        # 获取所有文件路径
        filenames = get_files(path)
        # 创建输出目录，返回是否存在和绝对路径
        exsit, abs_output_path = mkdir(output_path)
        # 遍历所有文件，并显示进度条
        for filename in simple_progress(filenames):
            # 判断文件的类型，对所有的 PPT 文件进行处理（PPT 文件以 ppt 或者 pptx 结尾）
            if check_suffix(filename, self.suffix_list):
                # 获取不带后缀的文件名，并生成新的 PDF 文件名，如：PPT素材1.pdf
                new_pdf_name = Path(filename).stem + '.pdf'  # PPT素材1.pdf
                # 生成完整的输出 PDF 文件路径
                output_pdf_filename = Path(abs_output_path).absolute() / new_pdf_name
                # 调用单个 PPT 转 PDF 的服务函数
                ppt2pdf_single(filename, str(output_pdf_filename))
                # time.sleep(3)

    def merge4ppt(self, input_path: str, output_path: str, output_name: str):
        """
        合并多个 PPT 文件为一个完整的 PPT 文件
        
        Args:
            input_path: PPT 文件所在的文件夹路径
            output_path: 合并后 PPT 文件的输出路径
            output_name: 合并后 PPT 文件的名称
        
        Returns:
            None
        """
        # if type(input_path) == list:
        #     ppt_file_list = [Path(ppt_file).absolute() for ppt_file in input_path]
        # else:
        # 将相对路径转换为绝对路径
        abs_input_path = Path(input_path).absolute()  # 相对路径→绝对路径
        # 获取文件夹下所有 PPT 文件列表
        ppt_file_list = get_files(abs_input_path)
        # 创建输出目录，返回是否存在和绝对路径
        exsit, abs_output_path = mkdir(output_path)
        # 生成完整的保存路径
        save_path = Path(abs_output_path) / output_name

        # 创建 PowerPoint 应用程序实例
        Application = win32com.client.gencache.EnsureDispatch(self.app)

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

    def ppt2img(self, input_path, output_path, merge=False):
        '''
        将 PPT 另存为图片格式，支持转换为单独图片或合并为长图
        
        Args:
            pptFullName: 要转换的 PPT 文件或文件夹路径
            pptName: 转换后的存放 JPG 文件的目录
            merge: 是否合并为长图，True 表示合并，False 表示不合并
        
        Returns:
            None
        '''

        # 获取所有 PPT 文件路径
        filenames = get_files(input_path)
        # 启动 PowerPoint 应用程序
        pptClient = win32com.client.Dispatch(self.app)
        # 设置为 0 表示后台运行，不显示，1 则显示
        pptClient.Visible = True
        # 遍历所有 PPT 文件
        for ppt_file in filenames:
            # 检查文件后缀是否为 ppt 或 pptx
            if check_suffix(ppt_file, self.suffix_list):
                # Python 路径操作模块 pathlib，看这篇就够了！https://zhuanlan.zhihu.com/p/475661402
                # 生成以 PPT 文件名为名称的输出目录
                output_dir = Path(output_path).absolute() / str(Path(ppt_file).stem)
                # 创建输出目录
                exsit, output_dir = mkdir(output_dir)
                # 打开 PPT 文件，WithWindow=False 表示不显示窗口，ReadOnly=1 表示只读
                ppt = pptClient.Presentations.Open(ppt_file, WithWindow=False, ReadOnly=1)
                # 将 PPT 另存为图片格式，17 表示 JPG 格式
                ppt.SaveAs(output_dir, self.ppt2img_type)

        # 退出 PowerPoint 应用程序
        pptClient.Quit()
        # 如果需要合并为长图
        if merge:
            # 获取输出目录的绝对路径
            merge_dir = Path(output_path).absolute()
            # 遍历输出目录下的所有子目录
            for dirpath, dirnames, filenames in os.walk(merge_dir):
                # print(dirpath, dirnames, filenames)
                # 遍历每个子目录
                for dirname in dirnames:
                    # 获取当前图片目录的完整路径
                    current_img_path = os.path.join(dirpath, dirname)
                    # 生成长图
                    self.generate_long_image(input_path=current_img_path,
                                             current_ppt_name=dirname,
                                             output_path=merge_dir,
                                             img_name=dirname + self.default_img_type)
                    # 删除暂存图片目录
                    shutil.rmtree(current_img_path)  # 删除暂存图片

    def generate_long_image(self, input_path: str, current_ppt_name, output_path, img_name='merge.jpg'):
        """
        将 PPT 的各个页面拼接成长图
        参考：https://blog.csdn.net/m0_51777056/article/details/130262561
        
        Args:
            input_path: 存放 PPT 各个页面图片的路径
            current_ppt_name: 当前 PPT 的名称
            output_path: 长图的输出路径
            img_name: 生成的长图文件名，默认为 'merge.jpg'
        
        Returns:
            None
        """
        # 获取图片列表
        img_list = []
        for imgs in os.listdir(input_path):
            img_list.append(os.path.join(input_path, imgs))

        # 将获取到 PPT 的页面进行排序，按照文件名长度排序
        ims_sort = sorted(img_list, key=lambda jpg: len(jpg))
        print(f"正在生成【PPT文件名为：{current_ppt_name}】的图片")

        # 取第一个图片的尺寸（宽度和高度）
        width, height = Image.open(img_list[0]).size  # 取第一个图片尺寸
        # 获取第一个图片的模式（RGB 或 RGBA）
        img_mode = Image.open(img_list[0]).mode
        # 创建同宽、n 倍高的空白图片，用于存放长图
        long_canvas = Image.new(img_mode, (width, height * len(img_list)))  # 创建同宽，n倍高的空白图片

        # 拼接图片，将每个页面图片依次粘贴到长画布上
        for i, image in enumerate(ims_sort):
            # 将图片粘贴到长画布的指定位置，从上到下排列
            long_canvas.paste(Image.open(image), box=(0, i * height))
        # 保存生成的长图
        long_canvas.save(os.path.join(output_path, img_name))  # 保存长图
