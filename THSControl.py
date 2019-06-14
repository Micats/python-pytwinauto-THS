#-*- coding:utf-8 -*-

"""
    时间：2019-6-3
    描述：对同花顺软件的控制：打开，关闭，登陆，切换登陆
"""

import os
import sys
import time
import threading
import psutil
import win32api
import pywinauto
import Structs
import Config




#打开功能
def OpenTHS():
    if IsProcess(Config.THSName):
        print("进程正在运行")
        return
    #r_v= os.system(Config.FullPath)  #这种方法启动程序，不会返回，需要等待进程结束才会继续执行
    #if r_v == 0:
    #    print("程序启动失败")
        
    win32api.ShellExecute(0, 'open', Config.FullPath, '','',1) #会立即返回，不用另开进程
    #print("程序启动成功")

#关闭功能
def CloseTHS():
    if IsProcess(Config.THSName):
        KillProcess(GetPid(Config.THSName))
        #处理弹窗
        time.sleep(Config.TIME_SAME_LAYER)
        #print(GetExitPopWin().print_ctrl_ids())
        #GetExitPopWin().Button1.click()
        GetExitPopWin()["同花顺Button3"].click()

#登录--分析平台
def LoginTHS(account,password):
    try:
        hwnd = pywinauto.findwindows.find_window(title="登录到行情主站",class_name=u'#32770')
    except:
        print("登录账号窗口未找到")
        return
    app = pywinauto.application.Application() #连接进程by标题名
    app.connect(title="登录到行情主站")
    win = app.window(handle = hwnd)
    #win[Switch["账户"]].set_focus()
    win[Structs.Switch["账户"]].set_edit_text(account)
    #time.sleep(Structs.TIME_SAME_LAYER)
    win[Structs.Switch["密码"]].set_focus()
    win[Structs.Switch["密码"]].set_edit_text(password)
    win[Structs.Switch["登录"]].click()

#切换账户--分析平台
def SwitchAccountTHS(account,password):
    if IsProcess(Config.THSName):
        KillProcess(GetPid(Config.THSName))
        #处理弹窗
        time.sleep(Config.TIME_SAME_LAYER)
        GetExitPopWin()["同花顺Button1"].click()
        time.sleep(Config.TIME_POP_WIN)
        try:
            hwnd = pywinauto.findwindows.find_window(title="登录到行情主站",class_name=u'#32770')
        except:
            print("切换账号窗口未找到")
            return
        app = pywinauto.application.Application() #连接进程by标题名
        app.connect(title="登录到行情主站")
        win = app.window(handle = hwnd)
        #win[Structs.Switch["账户"]].set_focus()
        win[Structs.Switch["账户"]].set_edit_text(account)
        #time.sleep(Structs.TIME_SAME_LAYER)
        win[Structs.Switch["密码"]].set_focus()
        win[Structs.Switch["密码"]].set_edit_text(password)
        win[Structs.Switch["登录"]].click()

#登录，选择已有资金账号--委托下单（账号类型）
def LoginCapitalAccount(type):
    dlg,app = GetLoginWin()
    try:        
        if dlg["确定"].exists(): #需要输入交易保护密码   此按钮只有交易密码界面有  
            dlg["AfxFrameOrView42s"].set_focus()
            dlg["AfxFrameOrView42s"].type_keys(Config.Account["pwd"]) #输入正确自动登录不同点击确定
            if dlg["确定"].exists():
                print("密码错误")
            return
    except:
        pass 
    dlg["Button3"].click()  #切换账户
    pop_hwnd = dlg.popup_window()
    win = app.window(handle = pop_hwnd)   
    win[type].click()
    dlg["登录"].click()
    #再次判断用不用填入保护密码
    dlg,app = GetLoginWin()
    try:        
        if dlg["确定"].exists(): #需要输入交易保护密码   此按钮只有交易密码界面有  
            dlg["AfxFrameOrView42s"].set_focus()
            dlg["AfxFrameOrView42s"].type_keys(Config.Account["pwd"]) #输入正确自动登录不同点击确定
            if dlg["确定"].exists():
                print("密码错误")
            return
    except:
        pass

    Structs.GL["CurrentAccount"] = GetCurAccount(type)
    

#切换登录，选择已有资金账号--委托下单（账号类型）
def SwitchCapitalAccount(type):
    try:
        mainApp=pywinauto.application.Application().connect(title=Config.WinTitle)
        mainHwnd = pywinauto.findwindows.find_window(title=Config.WinTitle) #找不到会抛出异常
        #if mainHwnd==0:
        #   print('请打开交易界面！')
        #   return
    except:
        print('请打开交易界面！')
        return
    mainWin=mainApp.window(handle = mainHwnd)
    mainWin["ToolbarWindow32"].button(1).click()  #["登录"]
    time.sleep(Config.TIME_POP_WIN)
    dlg,app = GetLoginWin()
    dlg["Button3"].click()  #切换账户   ["左上i角第一个键"]
    pop_hwnd = dlg.popup_window()
    win = app.window(handle = pop_hwnd)
    #print(win.print_ctrl_ids())
    win[type].click()
    dlg["登录"].click()
    #global CurrentAccount
    Structs.GL["CurrentAccount"] = GetCurAccount(type)
    #print(GL["CurrentAccount"])
    time.sleep(Config.TIME_POP_WIN)

#“网上交易系统”退出函数
def Exit():
    try:
        mainApp=pywinauto.application.Application().connect(title=Config.WinTitle)
        mainHwnd = pywinauto.findwindows.find_window(title=Config.WinTitle) #找不到会抛出异常
    except:
        #print('交易界面！')
        return
    mainWin=mainApp.window(handle = mainHwnd)
    mainWin["ToolbarWindow32"].button(0).click()  #["退出"]


#监视函数，监视软件是否运行
def Monitor():
    while True:
        if GetPid("xiadan.exe")==-1: 
            print("程序未运行,正在启动")
            pywinauto.application.Application().start(Config.XiaDanPath)
        else:
            print("程序正在运行")
            continue
        time.sleep(1)

#######################################以下是上述接口所需函数################################

#判断某个进程是否存在
def IsProcess(name):
    pl = psutil.pids()
    try:
        for pid in pl:
            if psutil.Process(pid).name() == name:
                return True
    except:
        pass
    return False

#得到某一进程的pid
def GetPid(name):
    pl = psutil.pids()
    for pid in pl:
        try:
            #val = psutil.Process(pid).name()
            if psutil.Process(pid).name() == name:       
                return pid
        except:
            #因为进程结束找不到而引起的错误
            continue
    return -1

#关闭某一进程
def KillProcess(pid):
    os.popen('taskkill.exe /pid:'+str(pid))


#得到连接的UIApp
def GetApp():
    app = pywinauto.application.Application() #连接进程by标题名
    app.connect(title="同花顺(v8.70.94) - 自选股")
    return app

#得到主界面，用来获取登陆，切换账户等功能
def GetMainHandle():
    return pywinauto.findwindows.find_window(title="同花顺(v8.70.94) - 自选股")
def GetMainWin():
    return GetApp().window(handle=GetMainHandle())   #主窗口

#得到点击“×”弹出的窗口
def GetExitPopWin():
    sub_hwnd = pywinauto.findwindows.find_windows(title = "同花顺",class_name=u'#32770', parent=GetMainHandle())
    if len(sub_hwnd) == 0:
        print("未获取到弹出窗口")
        return False
    return GetApp().window(handle = sub_hwnd[0])
    #pop_hwnd = GetMainWin().popup_window()
    #if pop_hwnd:
    #    return GetApp().window(handle=pop_hwnd)
    #return False

#得到委托下单登陆界面
def GetLoginWin():
    if GetPid("xiadan.exe")==-1: 
        app = pywinauto.application.Application().start(Config.XiaDanPath)
        time.sleep(Config.TIME_POP_WIN)
   #可能是登录界面，也可能已经登陆
    try:#已经登录
        hwnd = pywinauto.findwindows.find_window(title=Config.WinTitle) #找相关窗口
        #print("错误：已经登录，请使用切换登录功能")
        return
    except:#登录界面
        app = pywinauto.application.Application().connect(process = GetPid("xiadan.exe"))
        dlg = app.top_window()
        return dlg,app

#得到当前登陆账户
def GetCurAccount(val):
    for key in Config.Account.keys():
        if Config.Account[key]==val:
            return key
    return u"未定义"



if __name__ == "__main__":
    #OpenTHS()
    #time.sleep(15)
    #print("开始关闭")
    #CloseTHS()
    #SwitchAccount("11111","22222")
    #GetLoginWin()
    LoginCapitalAccount(Config.Account["icarusqaq"])
    #SwitchCapitalAccount(Config.Account["icarusqaq"])

    #subThread=threading.Thread(target=Monitor)
    #subThread.start()

    #dlg,app=GetLoginWin()
    #print(dlg.print_ctrl_ids())