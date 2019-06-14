#-*- coding:utf-8 -*-

"""
    时间：2019-6-13
    描述：配置文件
          资金账户，需要自动切换的，需要按时间顺序写进Account（dict）中
          例如 Account["icarusqaq"]="CPanel1"
          Account["用户账户"]="CPanel"+数字从1累加
"""



#交易界面顶层标题
WinTitle = u"网上股票交易系统5.0"
#同花顺软件所在的最终目录，不是根目录
THSPath="C:\\同花顺软件\\同花顺\\"
#同花顺软件名字
THSName = "hexin.exe"
#同花顺进程名字
ProcessName = "同花顺金融分析平台(32位)"
#全路径
FullPath = THSPath+THSName
#下单程序全路径,带有参数的启动，不知道参数的话，查看桌面快捷方式-委托交易
XiaDanPath="C:\\同花顺软件\\同花顺\\xiadan.exe ver=1 \"0711296a70d754486417e81e747c6d7a23d82c2c3de0ae1605a82d54ddb9facb64c7b32a56fd559b8ecc67abcd1cf2f8f0ce94d17b29d062e8eee500bb53ad78\""



#sleep的时间设置，不然操作太快有时候取不到句柄
TIME_SAME_LAYER = 0.5   #同一界面操作
TIME_DLG = 1          #快速非模态对话框
TIME_POP_WIN = 2      #弹出界面



#资金的登录账号界面
Account={}
Account["pwd"]="123456"       #六位保护密码
Account["icarusqaq"]="CPanel1"
Account["icarusqaq2"]="CPanel2"

