# 导入单元测试模块
import unittest
# 导入操作系统模块，用于路径和文件操作
import os
# 导入文件复制模块
import shutil

# 导入 poppt 的 API 接口
from poppt.api.ppt import *


class TestPPT(unittest.TestCase):
    """
    PPT 相关功能的单元测试类
    统一使用测试文件：./ppt/test_ppt/晚枫-1211重庆.pptx
    """

    @classmethod
    def setUpClass(cls):
        """
        测试类初始化，在所有测试用例执行前运行一次
        创建测试所需的输出目录
        """
        # 定义测试用的 PPT 文件路径
        cls.test_ppt = r'./ppt/test_ppt/晚枫-1211重庆.pptx'
        # 定义各个测试的输出目录
        cls.output_dirs = [
            r'./ppt/test_ppt/test_output_pdf',
            r'./ppt/test_ppt/test_output_img',
            r'./ppt/test_ppt/test_output_img_merge',
            r'./ppt/test_ppt/test_output_merge'
        ]

    def test_ppt2pdf(self):
        """
        测试 PPT 转 PDF 功能
        将单个 PPT 文件转换为 PDF 格式
        """
        output_path = r'./ppt/test_ppt/test_output_pdf'
        print(f"\n测试 PPT 转 PDF: {self.test_ppt} -> {output_path}")
        ppt2pdf(path=self.test_ppt, output_path=output_path)
        print("PDF 转换完成！")

    def test_ppt2img_single(self):
        """
        测试单个 PPT 转图片功能（每页独立保存）
        将 PPT 的每一页转换为独立的 JPG 图片
        """
        output_path = r'./ppt/test_ppt/test_output_img'
        print(f"\n测试 PPT 转图片（独立图片）: {self.test_ppt} -> {output_path}")
        ppt2img(input_path=self.test_ppt, output_path=output_path, merge=False)
        print("图片转换完成！")

    def test_ppt2img_merge(self):
        """
        测试 PPT 转图片功能（合并为长图）
        将 PPT 的所有页面合并为一张长图
        """
        output_path = r'./ppt/test_ppt/test_output_img_merge'
        print(f"\n测试 PPT 转长图: {self.test_ppt} -> {output_path}")
        ppt2img(input_path=self.test_ppt, output_path=output_path, merge=True)
        print("长图生成完成！")

    def test_merge4ppt(self):
        """
        测试合并 PPT 功能
        将文件夹中的所有 PPT 文件合并为一个 PPT
        注意：为了测试合并功能，先复制测试文件到临时目录
        """
        # 创建临时目录用于存放要合并的PPT文件
        temp_dir = r'./ppt/test_ppt/temp_merge_input'
        output_path = r'./ppt/test_ppt/test_output_merge'
        
        # 如果临时目录存在则先删除
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir, exist_ok=True)
        
        # 复制测试PPT文件两次，模拟多个PPT文件
        test_ppt_abs = os.path.abspath(self.test_ppt)
        shutil.copy(test_ppt_abs, os.path.join(temp_dir, '晚枫-1211重庆-副本1.pptx'))
        shutil.copy(test_ppt_abs, os.path.join(temp_dir, '晚枫-1211重庆-副本2.pptx'))
        
        print(f"\n测试合并 PPT: {temp_dir} -> {output_path}")
        merge4ppt(input_path=temp_dir, 
                  output_path=output_path, 
                  output_name='合并后的PPT.pptx')
        print("PPT 合并完成！")
        
        # 清理临时目录
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


if __name__ == '__main__':
    # 运行所有测试用例
    unittest.main(verbosity=2)
