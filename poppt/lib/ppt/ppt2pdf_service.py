"""
PPT 转 PDF 服务模块

开发说明：
1. 如何设置编辑器字体的大小?
   File(文件) -> Settings(设置) -> Editor(编辑器) -> Font(字体), 修改字体的大小
2. 注释: 代码的解释说明
"""

# 1). 导入需要的模块（打开应用程序的模块）
# 导入 Windows COM 组件，用于操作 PowerPoint
import win32com.client


def ppt2pdf_single(filename, output_filename):
    """
    将单个 PPT 文件导出为 PDF 格式
    
    Args:
        filename: PPT 文件的完整路径
        output_filename: 导出的 PDF 文件的完整路径
    
    Returns:
        None
    """
    # 2). 打开 PowerPoint 应用程序
    ppt_app = win32com.client.Dispatch('PowerPoint.Application')
    # 程序操作应用程序的过程是否可视化，True 表示可见
    ppt_app.Visible = True  # 程序操作应用程序的过程是否可视化

    # 3). 通过 PowerPoint 的应用程序打开指定的 PPT 文件
    # filename = "C:/Users/Administrator/Desktop/PPT办公自动化/ppt/PPT素材1.pptx"
    # output_filename = "C:/Users/Administrator/Desktop/PPT办公自动化/ppt/PPT素材1.pdf"
    # WithWindow=False 表示不显示窗口
    ppt = ppt_app.Presentations.Open(filename, WithWindow=False)

    # 4). 将打开的 PPT 另存为 PDF 文件
    # 17 是 PPT 转图片的数字，32 是 PPT 转 PDF 的数字
    ppt.SaveAs(output_filename, 32)
    # 退出 PowerPoint 应用程序
    ppt_app.Quit()
