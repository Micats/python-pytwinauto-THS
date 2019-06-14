#-*- coding:utf-8 -*-

"""
同花顺软件交易开发
时间：2019/5/7
开发环境：python 3.7(32bit)

注意:1.查询的文件，会保存在“文档”路径下，自动化更改不了，是一个下拉列表
     2.撤单无法精准撤一笔，原因是自定义控件，没找到合适的方法控制

"""

import os
import time
import sys
import pywinauto
import pandas as pd
import datetime
import time
import Structs
import THSControl
import Config

#得到当前日期
def GetCurDate():
    return time.strftime("%Y%m%d",time.localtime(time.time()))

"""
同花顺网上股票交易系统界面，自动化UI操作完成
"""
class THSSockDeal:
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        #self.__app_path=THSPath
        #print(GL["CurrentAccount"])
        self.__app = pywinauto.application.Application() #连接进程by标题名   
        while True:
            try:
                self.__app.connect(title=Config.WinTitle)
                self.__main_hwnd = pywinauto.findwindows.find_window(title=Config.WinTitle) #找相关窗口
                break
            except:
                continue
        #sub_hwnd = pywinauto.findwindows.find_windows(top_level_only=False, class_name=u'#32770', parent=main_hwnd)
        if self.__main_hwnd==0:
                print('请打开交易界面！')
                exit()       
        self.__main_window = self.__app.window(handle=self.__main_hwnd)   #主窗口

    def maxWindow(self):
        """
        最大化窗口
        """
        if self.__main_window.get_show_state() != 3:
            self.__main_window.maximize()
        self.__main_window.set_focus()

    def __closePopupWindow(self):
        """
        关闭一个弹窗。
        :return: 如果有弹出式对话框，返回True，否则返回False
        """
        popup_hwnd = self.__main_window.popup_window()
        if popup_hwnd:
            popup_window = self.__app.window(handle=popup_hwnd)
            popup_window.set_focus()
            #print(popup_window.static1.window_text())
            popup_window.Button.click()
            return True
        return False

    def __closePopupWindows(self):
        """
        关闭多个弹出窗口
        :return:
        """
        while self.__closePopupWindow():
            time.sleep(Config.TIME_DLG)

    def __buy(self,code,price,quantity):
        """ 买函数
        :param code: 代码， 字符串
        :param quantity: 数量， 字符串
        :param price: 买卖价格， 字符串
        """
        self.__main_window["SysTreeView32"].get_item([u"\\双向委托"]).click()
        self.__main_window["Edit1"].set_focus() #type_keys()函数是输入，不清空
        self.__main_window["Edit1"].set_edit_text(code)   
        self.__main_window["Edit2"].set_edit_text(price)
        self.__main_window["Edit3"].set_edit_text(quantity)
        self.__main_window[u"买入"].click()
        self.__closePopupWindows()
       
    def __sell(self,code,price,quantity):
        """ 卖函数
        :param code: 代码， 字符串
        :param quantity: 数量， 字符串
        :param price: 买卖价格， 字符串
        """
        self.__main_window["SysTreeView32"].get_item([u"\\双向委托"]).click()
        self.__main_window["Edit4"].set_focus() #type_keys()函数是输入，不清空
        self.__main_window["Edit4"].set_edit_text(code)   
        self.__main_window["Edit5"].set_edit_text(price)
        self.__main_window["Edit6"].set_edit_text(quantity)
        self.__main_window[u"卖出"].click()
        self.__closePopupWindows()

    def __inquiry_template_1(self,name):
        """
        查询模板，无任何输入，打开页面直接保存的类型
        """
        self.__closePopupWindows()
        try:
            self.__main_window["SysTreeView32"].get_item([u"\\查询[F4]",name]).click()
            self.__main_window["CVirtualGridCtrl"].click_input()
            self.__main_window["CVirtualGridCtrl"].type_keys('^s')
            time.sleep(Config.TIME_POP_WIN)
            popup_hwnd = self.__main_window.popup_window()
            if popup_hwnd == 0:
                return False
            hwnd = self.__app.window(handle=popup_hwnd)
            hwnd["Edit"].set_edit_text("%s_%s_%s.xls"%(Structs.GL["CurrentAccount"],name,GetCurDate()))
            hwnd[u"保存(&s)"].click()
            time.sleep(Config.TIME_SAME_LAYER)
            self.__closePopupWindow()
        except:
            print("查询失败")
            return

    def __inquiry_template_2(self,name,y_from,m_from,d_from,y_to,m_to,d_to,isInput=False,stock=""):
        """
        查询模版2，带日期，可以以股票代码分类查询
        """
        self.__closePopupWindows()
        try:
            self.__main_window.set_focus()
            if(name=="对账单"):
                name="对 账 单"
            self.__main_window["SysTreeView32"].get_item([u"\\查询[F4]",name]).click()
            #time.sleep(Config.TIME_SAME_LAYER)
            #self.__closePopupWindow()
            self.__main_window["CVirtualGridCtrl"].click_input()
            sub_hwnd = pywinauto.findwindows.find_windows(top_level_only=False, class_name=u'#32770',parent = self.__main_hwnd)[1]
            sub_win = self.__app.window(handle = sub_hwnd)
            sub_win["DateTimePicker"].set_time(year=y_from,month=m_from,day=d_from)
            sub_win["至DateTimePicker"].set_time(year=y_to,month=m_to,day=d_to)
            if isInput ==True:
                sub_win["Edit"].set_edit_text(stock)
            sub_win["确定"].click()
            time.sleep(Config.TIME_SAME_LAYER)
            self.__main_window["CVirtualGridCtrl"].type_keys('^s')
            time.sleep(Config.TIME_POP_WIN)
            popup_hwnd = self.__main_window.popup_window()
            if popup_hwnd == 0:
                return False
            hwnd = self.__app.window(handle=popup_hwnd)
            date_from = str(y_from)+str(m_from).zfill(2)+str(d_from).zfill(2)
            date_to = str(y_to)+str(m_to).zfill(2)+str(d_to).zfill(2)
            hwnd["Edit"].set_edit_text("%s_%s_%s_%s.xls"%(Structs.GL["CurrentAccount"],name,date_from,date_to))
            hwnd[u"保存(&s)"].click()
            time.sleep(Config.TIME_SAME_LAYER)
            self.__closePopupWindow()
        except:
            print("查询失败")
            return

    def order(self, code, quantity,price,direction):
        """
        下单函数
        :param code: 股票代码， 字符串
        :param direction: 买卖方向， 字符串
        :param quantity: 买卖数量， 字符串
        :param price: 买卖价格， 字符串
        """
        self.__closePopupWindows()
        if direction == 'B':
            self.__buy(code, quantity,price)
        if direction == 'S':
            self.__sell(code, quantity,price)
    
    def cancel_all(self):
        self.__main_window["SysTreeView32"].get_item([u"\\双向委托"]).click()
        self.__main_window["CVirtualGridCtrl"].click_input()
        self.__main_window[u"全撤(/)"].click()
        self.__closePopupWindows()

    def cancel_buy(self):
        self.__main_window["SysTreeView32"].get_item([u"\\双向委托"]).click()
        self.__main_window["CVirtualGridCtrl"].click_input()
        self.__main_window[u"撤买([)"].click()
        self.__closePopupWindows() 

    def cancel_sell(self):
        self.__main_window["SysTreeView32"].get_item([u"\\双向委托"]).click()
        self.__main_window["CVirtualGridCtrl"].click_input()
        self.__main_window[u"撤卖(])"].click()
        self.__closePopupWindows()

    def cancel_order(self,stock):
        self.__closePopupWindows()
        self.__main_window["SysTreeView32"].get_item([u"\\撤单[F3]"]).click()
        self.__main_window["Edit"].set_edit_text(stock)
        self.__main_window[u"查询代码"].click()
        #print(self.__main_window["CVirtualGridCtrl"].texts())
        self.cancel_all()
        self.__closePopupWindows()

    def inquiry_money(self):
        """
        可用资金查询,资金股票
        """
        self.__inquiry_template_1("资金股票")
        #data={} #文本信息
        #self.__main_window["SysTreeView32"].get_item([u"\\查询[F4]",u"\\资金股票"]).click()
        #for i in range(2,5):
        #    data[self.__main_window["Static%d"%(i)].texts()[0]]=self.__main_window["Static%d"%(i+3)].texts()[0]
        #for i in range(8,11):
        #    data[self.__main_window["Static%d"%(i)].texts()[0]]=self.__main_window["Static%d"%(i+3)].texts()[0]          
        #for i in range(14,20,2):
        #    data[self.__main_window["Static%d"%(i)].texts()[0]]=self.__main_window["Static%d"%(i+1)].texts()[0]
        #money = float( self.__main_window["Static5"].texts()[0])
      
        


    def inquiry_cur_deal(self):
        """
        当日成交查询
        """
        self.__inquiry_template_1("当日成交")

    def inquiry_cur_commit(self):
        """
        当日委托查询
        """
        self.__inquiry_template_1("当日委托")
        
    def inquiry_history_deal(self,fromyear=2019,frommonth=6,fromday=6,toyear=2019,tomonth=7,today=7):
        """
        历史成交查询
        """
        self.__inquiry_template_2("历史成交",fromyear,frommonth,fromday,toyear,tomonth,today)

    def inquiry_history_position(self,fromyear=2019,frommonth=6,fromday=6,toyear=2019,tomonth=7,today=7,stock=""):
        """
        历史持仓查询
        """
        self.__inquiry_template_2("历史持仓",fromyear,frommonth,fromday,toyear,tomonth,today,True,stock)
     
    def inquiry_history_commit(self,fromyear=2019,frommonth=6,fromday=6,toyear=2019,tomonth=7,today=7):
        """
        历史委托查询
        """
        self.__inquiry_template_2("历史委托",fromyear,frommonth,fromday,toyear,tomonth,today)

    def inquiry_capital_details(self,fromyear=2019,frommonth=6,fromday=6,toyear=2019,tomonth=7,today=7):
        """
        资金明细
        """
        self.__inquiry_template_2("资金明细",fromyear,frommonth,fromday,toyear,tomonth,today)

    def inquiry_acount_statement(self,fromyear=2019,frommonth=6,fromday=6,toyear=2019,tomonth=7,today=7):
        """
        对账单
        """
        self.__inquiry_template_2("对账单",fromyear,frommonth,fromday,toyear,tomonth,today)
        
if __name__ == "__main__":
    print("1")
    THSControl.LoginCapitalAccount(Config.Account["icarusqaq2"])
    app = THSSockDeal()
    app.order("000005","13.60","100",Structs.para["买入"])
    app.order("000001","13.60","100",Structs.para["卖出"])
    app.inquiry_money()
    app.inquiry_cur_deal()
    app.inquiry_cur_commit()
    app.inquiry_history_deal(2019,5,28,2019,5,29)
    app.inquiry_history_position(2019,5,28,2019,5,29,"000001")
    #app.inquiry_history_commit(2019,5,28,2019,5,29)
    app.inquiry_capital_details(2019,5,28,2019,5,29)
    app.inquiry_acount_statement(2019,5,28,2019,5,29)
    app.cancel_order("000001")

    THSControl.SwitchCapitalAccount(Config.Account["icarusqaq2"])
    app = THSSockDeal()
    app.order("000005","13.60","100",Structs.para["买入"])
    app.order("000001","13.60","100",Structs.para["卖出"])
    app.inquiry_money()
    app.inquiry_cur_deal()
    app.inquiry_cur_commit()
    app.inquiry_history_deal(2019,5,28,2019,5,29)
    app.inquiry_history_position(2019,5,28,2019,5,29,"000001")
    app.inquiry_capital_details(2019,5,28,2019,5,29)
    app.inquiry_acount_statement(2019,5,28,2019,5,29)
    app.cancel_order("000001")
    THSControl.Exit()

  
    
    
