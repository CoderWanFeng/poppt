# -*- coding: UTF-8 -*-
'''
@作者：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@读者群：http://www.python4office.cn/wechat-group/
@个人网站：www.python-office.com
@Date：2023/5/25 23:29 
@Description：合并 PPT 文件的示例代码
'''
# 导入 poppt 库
import poppt

# 设置要合并的 PPT 文件所在的文件夹路径
input_path = r"D:\workplace\code\github\poppt\dev\docs"
# 调用 merge4ppt 函数合并 PPT 文件
poppt.merge4ppt(input_path)
