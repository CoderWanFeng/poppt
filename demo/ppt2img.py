# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/yFcocJbfS9Hs375NhE8Gbw
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
                   merge=True)