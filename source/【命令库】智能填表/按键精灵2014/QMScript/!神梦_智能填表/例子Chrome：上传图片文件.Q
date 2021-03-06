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
MacroID=ed8d1758-2f55-4848-9d24-15d618881fcc
Description=例子Chrome：上传图片
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

//截图保存到C盘
Call Plugin.Pic.PrintScreen(0, 0, 100, 100, "C:\1.jpg")
Delay 2000  '强制等待 2 秒


//启动Chrome版谷歌内核
If LibE.神梦_智能填表.绑定("Chrome.exe") Then 

    '打开上传图片网址
    Call LibE.神梦_智能填表.网页_跳转("https://image.baidu.com/?fr=shitu")
    '判断网页是否加载完毕（超时 30 秒）
    Call LibE.神梦_智能填表.网页_是否就绪(30)


    //点击【识图一下】
    Call LibE.神梦_智能填表.元素_点击("tag:A&class:shitu-btn")
    Delay 2000  '强制等待 2 秒
    
    
    //选择要上传的文件（多个文件用“|”隔开）：
    Call LibE.神梦_智能填表.元素_上传文件("C:\1.jpg", "id:stfile")
    Delay 2000  '强制等待 2 秒
    
    MsgBox "脚本执行完毕！", 1+64+4096, "上传图片"
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If
