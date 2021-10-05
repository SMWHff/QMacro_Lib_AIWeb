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
MacroID=92125233-ef3b-44ff-b67f-713f0a9be412
Description=例子Chrome：跨域框架的表单测试
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


//启动Chrome浏览器
If LibE.神梦_智能填表.绑定("Chrome.exe") Then 
    Do 
        '打开189网址
        Call LibE.神梦_智能填表.网页_跳转("https://webmail30.189.cn/w2/")
        Delay 2000  '强制等待 2 秒
        
        '判断网页是否加载完毕加强版（关键字，0=文本判断；1=源码判断，超时 30 秒）
        If LibE.神梦_智能填表.网页_是否就绪加强("189邮箱", 1, 30) > 0 Then 
            TracePrint "网页已加载完毕！"
            Exit Do 
        Else 
            TracePrint "打开网页失败，正在重新加载。。。"
        End If
    Loop
	
	' 将网页滚动到右边
    Call LibE.神梦_智能填表.网页_置滚动条位置(574, 0)
    Delay 2000  '强制等待 2 秒
    
    ' 跨域框架输入帐号“SMWH”，没有指定框架会自动遍历所有框架，Chrome版谷歌内核必须加上元素特征
	[元素特征] = "INPUT[id='userName']"  
	Call LibE.神梦_智能填表.键盘_输入文本("SMWH")
    Delay 2000  '强制等待 2 秒
    
    ' 跨域框架输入密码“123456”，没有指定框架会自动遍历所有框架，Chrome版谷歌内核必须加上元素特征
    [元素特征] = "INPUT[id='password']"
    Call LibE.神梦_智能填表.键盘_输入文本("123456")
    Delay 2000  '强制等待 2 秒
    
    ' 跨框架勾选“一周自动登录”，没有指定框架会自动遍历所有框架
    Call LibE.神梦_智能填表.元素_置单复选框(True, "id:j-auto-login")
    Delay 2000  '强制等待 2 秒
    
    ' 跨框架点击按钮“登录”，可以指定框架索引，这样可以快速定位
    Call LibE.神梦_智能填表.元素_执行事件("click", "id:j-login&frame:0")
    Delay 3500  '强制等待 3.5 秒
    
    MsgBox "脚本执行完毕！", 1+64+4096, "跨域框架的表单测试"
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If
