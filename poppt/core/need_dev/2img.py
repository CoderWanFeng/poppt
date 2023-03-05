# -*- coding: UTF-8 -*-
'''
@Author  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@WeChat     ：CoderWanFeng
@Blog      ：www.python-office.com
@Date    ：2023/3/5 16:44
@Description     ：
'''

# https://blog.csdn.net/Selly166/article/details/126763130

import os
import win32com.client

'''
   支持转换为PNG、JPG，搜索后替换即可
'''
#JPG是17
ppSaveAsJPG = 17
#PNG是18
ppSaveAsPNG = 18

'''将PPT另存为图片格式
  arguments:
      pptFullName: 要转换的ppt文件，
      pptName：转换后的存放JPG文件的目录
      imgType: 图片类型
'''
def pptToImg(pptFullName, pptName, imgType):
    # 启动PPT
    pptClient = win32com.client.Dispatch('PowerPoint.Application')
    # 设置为0表示后台运行，不显示，1则显示
    pptClient.Visible = 1
    # 打开PPT文件
    ppt = pptClient.Presentations.Open(pptFullName)
    # 另存为图片
    ppt.SaveAs(pptName, imgType)
    # 退出
    pptClient.Quit()

'''
    多文件夹多图片文件重命名
'''
def renameImg(currentDir):
    #当前目录下，只获取所有的文件夹
    folders = [dI for dI in os.listdir(currentDir) if os.path.isdir(os.path.join(currentDir, dI))]
    i = 0;
    for folder in folders:
        #打开单个文件夹，获取文件列表
        fileList = os.listdir(folder)
        #重命名文件，规则1、2、3，开发时读取文件利于遍历
        i = 0;
        #遍历单个文件
        for file in fileList:
            #为了避免修改其他文件，只判断等于PNG、JPG的图
            fileFix = os.path.splitext(file)[-1]
            if fileFix.lower() == '.png' or fileFix == '.jpg':
                i += 1
                #完整路径img文件名 + 后缀，F:\my\projects\python\ppt2img\01_single_img/1.png
                imgFullName = os.path.join(currentDir, folder + '/' +  file)
                #完整路径img文件名不含后缀 F:\my\projects\python\ppt2img\01_single_img
                imgName = os.path.join(currentDir, folder)
                #重命名
                os.rename(imgFullName, imgName + '/' + (str(i) + fileFix.lower()))


if __name__ == '__main__':
    print("PPT转图片开始")
    # #获取当前路径
    currentDir = os.sys.path[0]
    #获取当前文件列表
    currentDirAllFiles = os.listdir(currentDir)
    #获取当前目录下所有后缀是ppt、pptx的文件，返回值是生成器对象（可迭代）
    currentDirPptFiles = (fns for fns in currentDirAllFiles if fns.endswith(('.ppt', '.pptx')))
    # 当前目录下所有的PPT文件名，和上述区别在于有无后缀名，返回值是生成器对象（可迭代）
    currentDirPptNames = (os.path.splitext(fns)[0] for fns in currentDirAllFiles if fns.endswith(('.ppt', '.pptx')))
    #fullFileName是文件名称 + 后缀01_single_img.pptx，fileName是文件名称不含后缀01_single_img
    for fullFileName, fileName in zip(currentDirPptFiles, currentDirPptNames):
        #完整路径ppt文件名 + 后缀，F:\my\projects\python\ppt2img\01_single_img.pptx
        pptFullName = os.path.join(currentDir, fullFileName)
        #完整路径PPT文件名 F:\my\projects\python\ppt2img\01_single_img
        pptName = os.path.join(currentDir, fileName)
        #需要创建一个与PPT同名的文件夹，判断下，如果不存在则创建
        if not os.path.exists(pptName):
            os.mkdir(pptName)
            #PPT转PNG
            pptToImg(pptFullName, pptName, ppSaveAsPNG)

           #PPT转JPEG
           # pptToImg(pptFullName, pptName, ppSaveAsJPG)

    print("PPT转图片完成")

    print("图片重命名开始")
    renameImg(currentDir)
    print("图片重命名完成，脚本执行结束")
