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
MacroID=7b490060-aa5d-41cc-949f-74d693ad1b78
Description=例子IE：遍历表格内容
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

//启动IE浏览器
If LibE.神梦_智能填表.绑定("IE.exe") Then 

    Do 
        '打开表格（table）测试网址
        Call LibE.神梦_智能填表.网页_跳转("http://www.w3school.com.cn/tiy/loadtext.asp?f=html_table_elements")
        Delay 2000  '强制等待 2 秒
        
        '判断网页是否加载完毕加强版（关键字，0=文本判断；1=源码判断，超时 30 秒）
        If LibE.神梦_智能填表.网页_是否就绪加强("这个单元包含一个表格", 1, 30) > 0 Then 
            TracePrint "网页已加载完毕！"
            Exit Do 
        Else 
            TracePrint "打开网页失败，正在重新加载。。。"
        End If
    Loop
    

    表总数 = LibE.神梦_智能填表.元素_取表格表总数("tag:table&index:0")
    For 表 = 0 To 表总数 - 1
        行总数 = LibE.神梦_智能填表.元素_取表格行总数(表, "tag:table&index:0")
        For 行 = 0 To 行总数 - 1
            列总数 = LibE.神梦_智能填表.元素_取表格列总数(表, 行, "tag:table&index:0")
            For 列 = 0 To 列总数 - 1
                TracePrint "Table"& 表 &"("& 行 & "," & 列 &") = " & LibE.神梦_智能填表.元素_取表格(表, 行, 列, "tag:table&index:0")
            Next 
        Next 
    Next

    MsgBox "脚本执行完毕！", 1+64+4096, "例子IE：遍历表格内容"
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If
