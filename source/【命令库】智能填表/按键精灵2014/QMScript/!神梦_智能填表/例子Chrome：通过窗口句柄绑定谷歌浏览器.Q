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
MacroID=73d907bc-5a58-4f51-a613-0fe6c93c0926
Description=例子Chrome：通过窗口句柄绑定谷歌浏览器
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


授权码 = "0123456789ABCDEF" '填入你自己的授权码

//记得要先初始化命令库，才能用命令库中的功能
If LibE.神梦_智能填表._初始化(授权码, "") = False  Then 
	MsgBox "出错，智能填表初始化失败！", 1 + 16 + 4096, "警告！"
	ExitScript
End If





//查找谷歌窗口句柄
Hwnd = Plugin.Window.Find("Chrome_WidgetWin_1", 0)
TracePrint "【谷歌窗口句柄】：" & Hwnd

//绑定谷歌浏览器窗口句柄
If LibE.神梦_智能填表.绑定(Hwnd) Then 

	'打开网址导航
	Call LibE.神梦_智能填表.网页_跳转("https://www.baidu.com/")
	Delay 2000  '强制等待 2 秒

	'输入“神梦填表”
	Call LibE.神梦_智能填表.元素_置输入框("神梦填表", "id:kw") 
	Delay 2000  '强制等待 2 秒
	
	'点击“百度一下”按钮
	Call LibE.神梦_智能填表.元素_点击("id:su")
	Delay 2000  '强制等待 2 秒
	
	MsgBox "脚本执行完毕！", 1+64+4096, "百度一下"
Else 
	MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
	ExitScript 
End If
