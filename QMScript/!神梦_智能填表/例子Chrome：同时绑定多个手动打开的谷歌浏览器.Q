[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=981ec4ea-905c-4b0d-b491-0101b9c0ea07
Description=例子Chrome：同时绑定多个手动打开的谷歌浏览器
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
'#================================================================
'#        【命令库】神梦_智能填表例子
'#----------------------------------------------------------------
'#        【作者】：神梦无痕
'#        【ＱＱ】：1042207232
'#        【Ｑ群】：624655641
'#        【更新】：2021-09-19
'#----------------------------------------------------------------
'#  插件说明：支持IE、Chrome内核网页填表（免费使用）
'#----------------------------------------------------------------
'#  神梦工具: http://pan.baidu.com/s/1dESHf8X
'#----------------------------------------------------------------
'#  特别声明：请勿用于非法用途！！（否则后果自负）
'#================================================================
'
'
/*【同时绑定多个已打开的谷歌浏览器的方法】*/
//第一步：
//    1、把桌面上的谷歌浏览器图标复制三个
//第二步：
//    1、将桌面上第一个谷歌浏览器图标，改名为【谷歌1】
//    	1) 右键图标，点击【属性】
//    	2)【目标】后面输入 --remote-debugging-port=9001 --user-data-dir=C:\SmSoft\Chrome_data\user_9001 --disable-web-security
//    	3)【注意】记得在 -- 就是每个双减号的前面都要加空格
//    	4)【注意】记得让 --remote-debugging-port 后面的数字和 \Chrome_data\user_ 后面的数字都为9001
//    	5) 点击【确定】
//    
//    2、将桌面上第二个谷歌浏览器图标，改名为【谷歌2】
//    	1) 右键图标，点击【属性】
//    	2)【目标】后面输入 --remote-debugging-port=9002 --user-data-dir=C:\SmSoft\Chrome_data\user_9002 --disable-web-security
//    	3)【注意】记得在 -- 就是每个双减号的前面都要加空格
//    	4)【注意】记得让 --remote-debugging-port 后面的数字和 \Chrome_data\user_ 后面的数字都为9002
//    	5) 点击【确定】
//    
//    3、将桌面上第三个谷歌浏览器图标，改名为【谷歌3】
//    	1) 右键图标，点击【属性】
//    	2)【目标】后面输入 --remote-debugging-port=9003 --user-data-dir=C:\SmSoft\Chrome_data\user_9003 --disable-web-security
//    	3)【注意】记得在 -- 就是每个双减号的前面都要加空格
//    	4)【注意】记得让 --remote-debugging-port 后面的数字和 \Chrome_data\user_ 后面的数字都为9003
//    	5) 点击【确定】
//第三步：
//    1、手动双击打开这三个谷歌浏览器图标
//    2、可以运行下面的代码了



授权码 = "0123456789ABCDEF" '填入你自己的授权码

//记得要先初始化命令库，才能用命令库中的功能
If LibE.神梦_智能填表._初始化(授权码, "") = False  Then 
	MsgBox "出错，智能填表初始化失败！", 1 + 16 + 4096, "警告！"
	ExitScript
End If


//绑定手动打开的【谷歌1】浏览器
BeginThread 多线程(9001, "https://www.baidu.com/") '参数直接写死，不要用变量

//绑定手动打开的【谷歌2】浏览器
BeginThread 多线程(9002, "https://www.baidu.com/") '参数直接写死，不要用变量

//绑定手动打开的【谷歌3】浏览器
BeginThread 多线程(9003, "https://www.baidu.com/") '参数直接写死，不要用变量


Sub 多线程(序号, 网址) 
    //启动谷歌浏览器   
    If LibE.神梦_智能填表.绑定("Chrome.exe" & ":" & 序号) Then
        
        '打开百度网址
        Call LibE.神梦_智能填表.网页_跳转(网址)
        Call LibE.神梦_智能填表.网页_是否就绪(30)
        Delay 2000  '强制等待 2 秒
           
        '输入“神梦填表”
        Call LibE.神梦_智能填表.元素_置输入框("神梦填表" & 序号, "id:kw")
        Delay 2000  '强制等待 2 秒
           
        '点击百度一下   
        Call LibE.神梦_智能填表.元素_点击("id:su")
        Delay 3500  '强制等待 3.5 秒
           
        MsgBox "脚本执行完毕！", 1+64+4096, "填表例子"
    Else
        MsgBox "出错，绑定浏览器（"& 序号 &"）失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
        EndScript
    End If
End Sub




