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
MacroID=974e093b-b822-4291-983d-62f5f0036566
Description=例子Chrome：处理对话框（弹窗）
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


//启动Chrome版谷歌内核
If LibE.神梦_智能填表.绑定("Chrome.exe") Then 

    Do 
        '打开对话框（alert）测试网址
        Call LibE.神梦_智能填表.网页_跳转("http://www.w3school.com.cn/tiy/t.asp?f=hdom_alert")
        Delay 2000  '强制等待 2 秒
        
        '判断网页是否加载完毕加强版（关键字，0=文本判断；1=源码判断，超时 30 秒）
        If LibE.神梦_智能填表.网页_是否就绪加强("显示消息框", 1, 30) > 0 Then 
            TracePrint "网页已加载完毕！"
            Exit Do 
        Else 
            TracePrint "打开网页失败，正在重新加载。。。"
        End If
    Loop 
    
    
    '====>>>【预处理弹窗】提前为即将弹出的窗口做准备，输入内容为空，确定按钮=true；取消按钮=false
    Call LibE.神梦_智能填表.网页_预处理弹窗("", True)
    
    '点击网页元素，弹出 alert 对话框
    Call LibE.神梦_智能填表.元素_点击("tag:input&value:显示消息框")
    
    '获取对话框（alert）上的内容：
    TracePrint LibE.神梦_智能填表.网页_取弹窗内容()  
    Delay 2000  '强制等待 2 秒
    '<<<====【预处理弹窗】与【取弹窗内容】 是成对出现，缺一不可！


    Do 
        '打开对话框（prompt）测试网址
        Call LibE.神梦_智能填表.网页_跳转("http://www.w3school.com.cn/tiy/t.asp?f=hdom_prompt")
        Delay 2000  '强制等待 2 秒
        
        '判断网页是否加载完毕加强版（关键字，0=文本判断；1=源码判断，超时 30 秒）
        If LibE.神梦_智能填表.网页_是否就绪加强("显示一个提示框", 1, 30) > 0 Then 
            TracePrint "网页已加载完毕！"
            Exit Do 
        Else 
            TracePrint "打开网页失败，正在重新加载。。。"
        End If
    Loop 

    '提前为即将弹出的窗口做准备，设置输入内容【神梦无痕】，点击【确定】按钮（true=确定；false=取消）
    Call LibE.神梦_智能填表.网页_预处理弹窗("神梦无痕", True)
    
    '点击网页元素，弹出 prompt 对话框
    Call LibE.神梦_智能填表.元素_点击("tag:input&value:显示一个提示框")
    
    '获取对话框（prompt）上的提示内容：
    TracePrint LibE.神梦_智能填表.网页_取弹窗内容()  
    Delay 2000  '强制等待 2 秒
    
    MsgBox "脚本执行完毕！", 1+64+4096, "处理对话框（弹窗）"
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If

