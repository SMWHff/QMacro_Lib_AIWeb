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
MacroID=026e115a-cc2a-4f72-9afd-4ec7e291450c
Description=例子2：解决授权后的小精灵无法保存QUI界面配置问题
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



//授权后的小精灵无法保存QUI界面配置，需要加上如下代码
'加载配置文件
Event Form1.Load
    Call Plugin.File.CopyFile("C:\uservar.ini", GetExeDir() & "uservar.ini" )
End Event

'保存配置文件
Event Form1.UnLoad
    Form1.SaveSetting : Delay 1000
    Call Plugin.File.CopyFile(GetExeDir() & "uservar.ini", "C:\uservar.ini")
End Event
