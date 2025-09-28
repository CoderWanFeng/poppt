# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@读者群     ：http://www.python4office.cn/wechat-group/
@个人网站 ：www.python-office.com
@Date    ：2023/6/8 23:26 
@Description     ：
'''
# import poppt
#
ppt_path = r'D:\workplace\code\github\poppt\demo\out\a\我的介绍.pptx'
out_dir = r'D:\workplace\code\github\poppt\demo\out\a\b\c'

# poppt.ppt2img(ppt_path, out_dir, merge=True)

import office

office.ppt.ppt2img(input_path=ppt_path,
                   output_path=out_dir,
                   merge=False)
