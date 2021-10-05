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
MacroID=6db869d0-ac26-4935-ba30-88b483b20c2e
Description=例子Chrome：超人打码测试
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


打码平台 = "CR" '（CR=超人打码，LZ=联众打码，CJY=超级鹰打码）
打码账号 = InputBox("请输入打码账号：", "例子Chrome：超人打码测试", Plugin.File.ReadINI("神梦抓抓工具", "打码账号" & 2, "C:\神梦辅助\Config.ini"))
打码密码 = InputBox("请输入打码密码：", "例子Chrome：超人打码测试", Plugin.File.ReadINI("神梦抓抓工具", "打码密码" & 2, "C:\神梦辅助\Config.ini"))

授权码 = "0123456789ABCDEF" '填入你自己的授权码

//记得要先初始化命令库，才能用命令库中的功能
If LibE.神梦_智能填表._初始化(授权码, "") = False  Then 
    MsgBox "出错，智能填表初始化失败！", 1 + 16 + 4096, "警告！"
    ExitScript 
End If

'查询打码余额
返回值 = LibE.神梦_智能填表.接码_图形登录(打码平台,打码账号,打码密码)
TracePrint "【打码余额】：" & 返回值


//启动谷歌浏览器
If LibE.神梦_智能填表.绑定("Chrome.exe") Then 
    
    Do 
        //跳转到测试页面
        Call LibE.神梦_智能填表.网页_跳转("http://cp.3v.do/host/index.htm")
        Delay 2000  '强制等待 2 秒
        
        '判断网页是否加载完毕加强版（关键字，0=文本判断；1=源码判断，超时 30 秒）
        If LibE.神梦_智能填表.网页_是否就绪加强("虚拟主机控制面板", 1, 30) > 0 Then 
            TracePrint "网页已加载完毕！"
            Exit Do 
        Else 
            TracePrint "打开网页失败，正在重新加载。。。"
        End If
    Loop 
    
    //这里测试用的是QQ超人打码
    返回值 = LibE.神梦_智能填表.接码_图形验证(打码平台, "tag:IMG&index:1", 打码账号, 打码密码, "57716")
    If Left(返回值, 1) <> "-" Then 
        TracePrint "【验证码】：" & 返回值
        '输入验证码
        Call LibE.神梦_智能填表.元素_置输入框(返回值, "tag:input&name:code")
    Else
        If 返回值 = "-1" Then 
            MsgBox "账号或密码不正确", 16 + 4096, "打码失败！"
        ElseIf 返回值 = "-2" Then 
            MsgBox "用户不存在", 16 + 4096, "打码失败！"
        ElseIf 返回值 = "-3" Then 
            MsgBox "余额不足", 16 + 4096, "打码失败！"
        ElseIf 返回值 = "-4" Then 
            MsgBox "验证码类型错误", 16 + 4096, "打码失败！"
        Else 
            MsgBox 返回值, 16 + 4096, "打码失败！"
        End If
        ExitScript 
    End If
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If

'打码报错
TracePrint "【打码报错】：" & LibE.神梦_智能填表.接码_图形报错(打码平台, 打码账号, 打码密码)

MsgBox "脚本执行完毕！", 1+64+4096, "超人打码测试"

