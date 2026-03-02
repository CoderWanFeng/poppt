# -*- coding: UTF-8 -*-
'''
@作者：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@读者群：https://www.python4office.cn/wechat-group/
@个人网站：www.python-office.com
@Date：2023/3/5 16:44
@Description：PPT 转图片的开发测试代码（待开发完善）
@参考：https://blog.csdn.net/Selly166/article/details/126763130
'''

# https://blog.csdn.net/Selly166/article/details/126763130

# 导入系统操作相关模块
import os
# 导入 Windows COM 组件，用于操作 PowerPoint
import win32com.client

'''
   支持转换为 PNG、JPG，搜索后替换即可
'''
# JPG 对应的数字是 17
ppSaveAsJPG = 17
# PNG 对应的数字是 18
ppSaveAsPNG = 18

'''
将 PPT 另存为图片格式

Args:
    pptFullName: 要转换的 PPT 文件完整路径
    pptName: 转换后的存放 JPG 文件的目录
    imgType: 图片类型，17 为 JPG，18 为 PNG
'''
def pptToImg(pptFullName, pptName, imgType):
    # 启动 PowerPoint 应用程序
    pptClient = win32com.client.Dispatch('PowerPoint.Application')
    # 设置为 0 表示后台运行，不显示，1 则显示
    pptClient.Visible = 1
    # 打开 PPT 文件
    ppt = pptClient.Presentations.Open(pptFullName)
    # 将 PPT 另存为指定的图片格式
    ppt.SaveAs(pptName, imgType)
    # 退出 PowerPoint 应用程序
    pptClient.Quit()

'''
    多文件夹多图片文件重命名
'''
def renameImg(currentDir):
    # 当前目录下，只获取所有的文件夹
    folders = [dI for dI in os.listdir(currentDir) if os.path.isdir(os.path.join(currentDir, dI))]
    i = 0;
    # 遍历每个文件夹
    for folder in folders:
        # 打开单个文件夹，获取文件列表
        fileList = os.listdir(folder)
        # 重命名文件，规则 1、2、3，开发时读取文件利于遍历
        i = 0;
        # 遍历单个文件
        for file in fileList:
            # 为了避免修改其他文件，只判断等于 PNG、JPG 的图片
            fileFix = os.path.splitext(file)[-1]
            # 检查文件后缀名是否为 png 或 jpg（忽略大小写）
            if fileFix.lower() == '.png' or fileFix == '.jpg':
                i += 1
                # 完整路径 img 文件名 + 后缀，F:\my\projects\python\ppt2img\01_single_img/1.png
                imgFullName = os.path.join(currentDir, folder + '/' +  file)
                # 完整路径 img 文件名不含后缀 F:\my\projects\python\ppt2img\01_single_img
                imgName = os.path.join(currentDir, folder)
                # 重命名文件，格式为：数字 + 后缀名
                os.rename(imgFullName, imgName + '/' + (str(i) + fileFix.lower()))


if __name__ == '__main__':
    print("PPT转图片开始")
    # # 获取当前路径
    currentDir = os.sys.path[0]
    # 获取当前文件列表
    currentDirAllFiles = os.listdir(currentDir)
    # 获取当前目录下所有后缀是 ppt、pptx 的文件，返回值是生成器对象（可迭代）
    currentDirPptFiles = (fns for fns in currentDirAllFiles if fns.endswith(('.ppt', '.pptx')))
    # 当前目录下所有的 PPT 文件名，和上述区别在于有无后缀名，返回值是生成器对象（可迭代）
    currentDirPptNames = (os.path.splitext(fns)[0] for fns in currentDirAllFiles if fns.endswith(('.ppt', '.pptx')))
    # fullFileName 是文件名称 + 后缀 01_single_img.pptx，fileName 是文件名称不含后缀 01_single_img
    for fullFileName, fileName in zip(currentDirPptFiles, currentDirPptNames):
        # 完整路径 ppt 文件名 + 后缀，F:\my\projects\python\ppt2img\01_single_img.pptx
        pptFullName = os.path.join(currentDir, fullFileName)
        # 完整路径 PPT 文件名 F:\my\projects\python\ppt2img\01_single_img
        pptName = os.path.join(currentDir, fileName)
        # 需要创建一个与 PPT 同名的文件夹，判断下，如果不存在则创建
        if not os.path.exists(pptName):
            os.mkdir(pptName)
            # PPT 转 PNG
            pptToImg(pptFullName, pptName, ppSaveAsPNG)

           # PPT 转 JPEG
           # pptToImg(pptFullName, pptName, ppSaveAsJPG)

    print("PPT转图片完成")

    print("图片重命名开始")
    # 对生成的图片进行重命名
    renameImg(currentDir)
    print("图片重命名完成，脚本执行结束")
