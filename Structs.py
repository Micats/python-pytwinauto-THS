#-*- coding:utf-8 -*-

"""
    时间：2019-6-13
    描述：一些界面的控件dict及信息
    备注：CurrentAccount与文件名相关联，非自动化登录是赋值不了的，会默认为“未定义”
"""

#order的买卖方向
para={}
para["买入"]="B"
para["卖出"]="S"

#切换账号
Switch={}
Switch["账户"]=u"Edit1"
Switch["密码"]=u"Edit2"
Switch["自动登录"]=u"Button1"
Switch["保存密码登录"]=u"Button2"
Switch["登录"]=u"Button3"

#当前登录的账号,全局变量
GL={}
GL["CurrentAccount"]="未定义"

