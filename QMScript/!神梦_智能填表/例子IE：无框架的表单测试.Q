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
MacroID=6fedbb02-6952-431c-8931-54d2657fa1f3
Description=例子IE：无框架的表单测试
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
        '打开一个无框架的测试网址 
        Call LibE.神梦_智能填表.网页_跳转("http://www.vrbrothers.com/cn/wqm/demo/pages/Demo-ComplexForm.aspx")
        Delay 2000  '强制等待 2 秒
        
        '判断网页是否加载完毕加强版（关键字，0=文本判断；1=源码判断，超时 30 秒）
        If LibE.神梦_智能填表.网页_是否就绪加强("WQM测试表单", 1, 30) > 0 Then 
            TracePrint "网页已加载完毕！"
            Exit Do 
        Else 
            TracePrint "打开网页失败，正在重新加载。。。"
        End If
    Loop 
    
    '昵称
    Call LibE.神梦_智能填表.元素_置输入框("神梦无痕", "id:ctl00_mainContent_tbUsername")
    Delay 2000  '强制等待 2 秒
    
    
    '密码 
    Call LibE.神梦_智能填表.元素_置输入框("abc123456", "id:ctl00_mainContent_tbPassword")
    Delay 2000  '强制等待 2 秒
    
    
    '电子邮箱 
    Call LibE.神梦_智能填表.元素_置输入框("1042207232@qq.com", "id:ctl00_mainContent_tbEMail")
    Delay 2000  '强制等待 2 秒
    
    
    '性别：男 
    Call LibE.神梦_智能填表.元素_置单复选框(True, "id:ctl00_mainContent_rblGender_0")
    Delay 2000  '强制等待 2 秒
    
    
    '省份：广东
    Call LibE.神梦_智能填表.元素_置下拉框("广东", "id:ctl00_mainContent_ddlProvince")
    Delay 2000  '强制等待 2 秒
    
    
    '城市：深圳
    Call LibE.神梦_智能填表.元素_置下拉框("深圳", "id:ctl00_mainContent_ddlCity")
    Delay 2000  '强制等待 2 秒
    
    
    '求职意向：①销售； ④计算机软、硬件/互联网/IT
    Call LibE.神梦_智能填表.元素_置下拉框("销售%计算机软、硬件/互联网/IT", "id:ctl00_mainContent_lbObjectives")
    Delay 2000  '强制等待 2 秒
    
    
    '爱好：音乐，电影
    Call LibE.神梦_智能填表.元素_置单复选框(True, "id:ctl00_mainContent_cblHobies_0")
    Call LibE.神梦_智能填表.元素_置单复选框(True, "id:ctl00_mainContent_cblHobies_2")
    Delay 2000  '强制等待 2 秒
    
    
    '自我评价（换行符用 vbCrLf 表示）
    文本内容 = "作者：神梦无痕" & vbCrLf
    文本内容 = 文本内容 & "ＱＱ：1042207232" & vbCrLf
    文本内容 = 文本内容 & "Ｑ群：624655641"
    Call LibE.神梦_智能填表.元素_置输入框(文本内容, "id:ctl00_mainContent_tbSelfAssement")
    Delay 2000  '强制等待 2 秒
    
    
    '我已阅读并接受注册协议
    Call LibE.神梦_智能填表.元素_置单复选框(True, "id:ctl00_mainContent_cbAcceptTerms")
    Delay 2000  '强制等待 2 秒
    
    
    '提交
    Call LibE.神梦_智能填表.元素_点击("id:ctl00_mainContent_btnSubmit")
    Delay 2000  '强制等待 2 秒
    
    
    '判断网页是否加载完毕（超时设置 30 秒）
    Call LibE.神梦_智能填表.网页_是否就绪(30)
    
    '获取结果
    结果 = ""
    行数 = LibE.神梦_智能填表.元素_取表格("", - 1 , "", "tag:table&index:0")'获取表格行数
    For h = 0 To 行数 - 1
        列数 = LibE.神梦_智能填表.元素_取表格("", h, - 1 , "tag:table&index:0")'获取表格 h 行的列数
        Str = "" 
        For l = 0 To 列数 - 1
            s = LibE.神梦_智能填表.元素_取表格("", h, l, "tag:table&index:0")
            Str = Str & " " & s
        Next
        结果 = 结果 & Str & vbCrLf 
        TracePrint Str
    Next
    
    MsgBox 结果, 1+64+4096, "无框架的表单测试"
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If
