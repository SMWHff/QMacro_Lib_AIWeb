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
MacroID=498611b9-5a0b-4a03-9f89-f80c21a1e654
Description=例子Chrome：将谷歌里获取的Cookie转为Set-Cookie响应头格式
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


授权码 = "0123456789ABCDEF"  '填入授权码

//记得要先初始化命令库，才能用命令库中的功能
If LibE.神梦_智能填表._初始化(授权码, "") = False  Then 
    MsgBox "出错，智能填表初始化失败！", 1 + 16 + 4096, "警告！"
    ExitScript 
End If


//启动谷歌浏览器
If LibE.神梦_智能填表.绑定("Chrome.exe") Then 

    '打开百度网址
    Call LibE.神梦_智能填表.网页_跳转("https://www.baidu.com/")
    Delay 2000
    
    '谷歌获取的Cookie是JSON格式的
    Cookies =  LibE.神梦_智能填表.网页_取Cookie()
    
    'JSON格式转为 set-Cookie 响应头格式
    Call Plugin.SMWH.GetJSON(Cookies, "cookies", count)
    RetCookies = ""
    For i = 0 To count - 1
       	name = Plugin.SMWH.GetJSON(Cookies, "cookies[" & i & "].name")
       	value = Plugin.SMWH.GetJSON(Cookies, "cookies[" & i & "].value")
       	Path = Plugin.SMWH.GetJSON(Cookies, "cookies[" & i & "].path")
       	Domain = Plugin.SMWH.GetJSON(Cookies, "cookies[" & i & "].domain")
       	expires = Plugin.SMWH.GetJSON(Cookies, "cookies[" & i & "].expires")
       	GMT = Plugin.SMWH.JS_Run(1, "new Date("& expires &").toGMTString()")
       	Secure = Plugin.SMWH.GetJSON(Cookies, "cookies[" & i & "].sourceScheme")
       	HttpOnly = Plugin.SMWH.GetJSON(Cookies, "cookies[" & i & "].httpOnly")
       	If Secure <> "" Then Secure = "Secure;"
       	If HttpOnly Then 
            HttpOnly = "HttpOnly;"
       	Else 
            HttpOnly = ""
       	End If
       	TracePrint HttpOnly
       	RetCookies = RetCookies & " " &_
                    name & "=" & value & ";"&_
                    "expires=" & GMT & ";" &_
                    "path="& Path & ";" &_
                    "domain="& Domain & ";" &_
                    Secure & HttpOnly
    Next
    TracePrint RetCookies
    
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If