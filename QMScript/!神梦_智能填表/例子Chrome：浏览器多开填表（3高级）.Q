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
MacroID=f973dd35-7650-4db8-add7-ce807d168231
Description=例子Chrome：浏览器多开填表（3高级）
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=1
QUI=Form1
[Relative]
SetupOCXFile=
[Comment]

[UIPackage]
UEsDBBQAAgAIAIpxM1OAne7GoAMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAdVRUdhVUVHYVVFR2HtWUtPU0EUPrfl0QeWlkdBHnIBUZewcGdMgEhigo/QqgmJJhesSiyFtDVBf4V/wbUxYkKMv8C4MNGdW1y4dmPYSf3m3rl0aCDM45ZrEz5ymNDeud/MmfMcouTh+7fM7psPQz+pAdcpSvu1OHUIn1lcXKSJIvzv/Vqt5n9cO0NL4S+kjZ8hO+t2SCckDolBEpBzkCSkC5Lyjp66IRk+b/9MjS2LJdrAT5VsukEljGV6SSrIwmL8d1knPGv/7n0w+u6HFSVuYMA9uknTpI8YWZbPHznmmfQXj9fnF7+bpzwtGqwgwfkHuO+o8ueh700q0G1yaB2jKnoo4vJneOyVndfGR99//xc5w+liYaO8buB+xObPGMy3yAym802xs/31l8n+FVy2KfN14049/kUO1XyyZ+bHSpZ1HFqhouYaUoh/SV6nyPJHhPh3n9aoAtFdQVpj/1GBfwmsDj2mO9BDUTH3+vG/i79Tlr9N4M+Dfwt53+D8LdX9twvnP4/qowgpa66kT4O/Q9h/nX8O57BKz9X53Xq5W0H/nQ3nv2qof1afxxT4Y+TV8CIewin2qPXQGTL/3tE6d9PSrZydc0oVO1corz1pEv/U20dG85cRecqIPwX81gGLv/3cp2TtLy7Y/yJ8T93rDtt/mr9Plj9xpP/nsI5n0INaNhyE//cSi8Py/EmBfxbV/6abA8qQEtagloWy2H+cvL5clr9L4F/A3ktm8Uc5/7F4de1g/0Vk36dYA9N7SSf+K/OnDp2/A/1XsQamhzn3HPxMULeN43EF+s9S/d5Lhr9b4J9z3+/5nwzfEfZn9SnypwV+0/hhWn+aYtFZKRRN5xsU8K3ev+Rh+VW39tWxPjf+KPtfRrC/u+DfAnsRVbgOUhr8PQ38rP56gX3r3MIMavCL9+Xe/ZO+ASZgQT730DHP+PdOzbx/ukTy+YflS7/+D6P/ZPHS/79GEP2njXFMYf/9FGj/aY0Ty0Py/FkKtv8cdv1Ann+Awu0/2VqTB/53+v3neQq2/xzBeFFB/0MC//aOvTAzbX/+8/G9lv7Drh8+Qa5Cli3Wy6gj7PwbDZk/iPuHCYyjCvY3TPX7hwD6T+X4P0LB9p9T3Kdl+Ucp0P7TjT+XFfgvULD9J3vfpAL/GIXbf9oUbv85TmL+8/RewSk4bg2qtgad+nNC4N9tmPOaB0QLbn0a91emCDt+tzpm0flU4YM5eMErvfpf2f4mBfsLAqr8QaOV+f8BUEsBAhcLFAACAAgAinEzU4Cd7sagAwAACCYAAAkACQAAAAAAAAAAAACAAAAAAFVJUGFja2FnZVVUBQAHVUVHYVBLBQYAAAAAAQABAEAAAADYAwAAAAA=


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


Global 多开序号, 多开网址, 多开无头, 多开无痕, 多开IP, 多开UA, 多开Cookie, 多开原始参数
Global 入栈锁, 阻塞锁_入栈, 阻塞锁_出栈, 自旋锁, 进度锁, 总任务数

授权码 = "0123456789ABCDEF" '填入你自己的授权码

//记得要先初始化命令库，才能用命令库中的功能
If libE.神梦_智能填表._初始化(授权码, "") = False  Then
    MsgBox "出错，智能填表初始化失败！", 1 + 16 + 4096, "警告！"  
    ExitScript    
End If

'创建多线程锁
阻塞锁_入栈 = 阻塞锁创建()
阻塞锁_出栈 = 阻塞锁创建()
自旋锁 = 自旋锁创建()


'结束残留的谷歌进程
Call lib.系统.结束进程("Chrome.exe")


//'设置同时打开浏览器窗口的数量
并发上限 = 5
总任务数 = 10

For q = 1 To 总任务数
    ReDim 线程ID数组(并发上限)
    For i = 1 To 并发上限
        入栈锁 = True'开始入栈，打开入栈锁开关
        
        //记住，下面 8 个变量要传入多线程，必须用 Global 定义成全局变量
        多开序号 = i
        多开网址 = "https://www.baidu.com/s?wd=" & i
        多开无头 = False  'True=隐藏窗口；False=显示窗口
        多开无痕 = True   'True=开启隐身；False=关闭隐身
        多开IP = ""       '代理IP格式为【127.0.0.1:80】
        多开UA = "Opera/9.80 (Windows NT 6.1; U; zh-cn) Presto/2.9.168 Version/11.50"
        多开Cookie = ""
        多开原始参数 = "--disable-popup-blocking||--disable-desktop-notifications" '多个参数用“||”隔开
        
        线程ID数组(i) = BeginThread(多线程(多开序号, 多开网址, 多开无头, 多开无痕, 多开IP, 多开UA, 多开Cookie, 多开原始参数))
        
        '判断入栈锁状态，True=循环等待；False=跳出循环
        While(入栈锁 = True)
            Delay 500
        Wend
    Next
    
    
    
    '***********************************（重要！勿动）调节多线程的入栈出栈*************************************************
    '所有线程的浏览器都打开了，关闭【阻塞锁_入栈】开关
    Call 阻塞锁退出(阻塞锁_入栈)
    
    '判断所有线程任务是否完成（没完成继续等待，完成了跳出循环）
    While(进度锁 < 并发上限)
        Delay 3000
    Wend
    
    '所有线程的浏览器都关闭了，关闭【阻塞锁_出栈】开关
    Call 阻塞锁退出(阻塞锁_出栈)
    
    '判断所有线程是否停止
    Do
        For n = 1 To 并发上限
            If 取线程运行状态(线程ID数组(n)) = 1 Then Goto Do循环尾
            Delay 100
        Next
        Exit Do 
        Rem Do循环尾
    Loop 
    
    '所有线程完成任务，进度归零，进入下一轮
    进度锁 = 0
    
    '【阻塞锁_入栈】已经失效了，需要重新创建
    阻塞锁_入栈 = 阻塞锁创建()
    '***********************************（重要！勿动）调节多线程的入栈出栈*************************************************
Next








Sub 多线程(浏览器序号, 浏览器网址, 是否无头模式, 是否无痕模式, 设置代理IP, 设置UserAgent, 设置Cookies, 原始参数)
    '***********************************（重要！勿动）启动并绑定浏览器*************************************************
    //获取传入参数，重新赋值新变量，以免被其他线程影响 
    'Dim 序号  //不要对【序号】进行定义，它在后面还有重要作用
    Dim 网址, 扩展参数, x, y
    
    序号 = 浏览器序号
    网址 = 浏览器网址
    
    '设置浏览器参数
    扩展参数 = Array(是否无头模式, 是否无痕模式, 设置代理IP, 设置UserAgent, 设置Cookies, 原始参数)
    
    If libE.神梦_智能填表.绑定多开("Chrome.exe", 序号, "", 扩展参数) = False Then 
        MsgBox "【多开失败】：" & 序号, 16 + 4096, "警告！"
        Exit Sub 
    End If
    TracePrint "【多开序号】：" & 序号
    Randomize
    x = 序号-1 : y = 0
    If 序号 > 5 Then x = ((序号-1) Mod 5) : y = Fix((序号-1) / 5)
    x = x * 250
    y = y * 200
    Call libE.神梦_智能填表.网页_置浏览器位置(x, y)
    
    入栈锁 = False'关闭入栈锁（打开浏览器时，需要利用入栈锁排队一个一个打开，不能同时开，否则会冲突）
    
    '为了增加稳定性，在这里加阻塞，等待所有线程的浏览器都打开完毕再继续执行
    Call 阻塞锁进入(阻塞锁_入栈)
    '***********************************（重要！勿动）启动并绑定浏览器*************************************************
    
    
	
    
    
    //=============================【这里面写你自己的代码】=============================
    Dim n
    Call libE.神梦_智能填表.网页_跳转(网址)
    Call libE.神梦_智能填表.网页_是否就绪(30)

    TracePrint "【多开序号："& 序号 &"】" & libE.神梦_智能填表.网页_取标题()
    TracePrint "【多开序号：" & 序号 & "】" & libE.神梦_智能填表.网页_取网址()
    
    For n = 0 To 100
        'TracePrint 序号 & "_" & n
        Call libE.神梦_智能填表.元素_置输入框(序号 & "_" & n, "#kw")
        Delay 100  '强制等待 0.1 秒
    Next
End Sub


//以下是只对【多线程停止时】产生的事件      
Sub OnThreadExit()
    '***********************************（重要！勿动）关闭浏览器*************************************************
    If 序号 > 0 Then 
        Call libE.神梦_智能填表.关闭浏览器()
        
        TracePrint "【" & 序号 & "】正在等待出栈。。。"
        Call 自旋锁进入(自旋锁) : 进度锁 = 进度锁 + 1 : Call 自旋锁退出(自旋锁)
        Call 阻塞锁进入(阻塞锁_出栈)
        
        TracePrint "【" & 序号 & "】结束！"
        序号 = 0
    End If
    TracePrint TypeName(oSMWH)
    Set oSMWH = Nothing 
    '***********************************（重要！勿动）关闭浏览器*************************************************
End Sub 










//===============================【子线程·不懂勿动】=====================================================
'
'--------------------------------[取线程状态]--------------------------------
Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Function 取线程运行状态(线程ID)
    Dim State
    State = 0
    If IsNumeric(CStr(线程ID)) Then State = Abs(WaitForSingleObject(OpenThread(&H1F0FFF, 0, CLng(线程ID)), 0) = 258)
    取线程运行状态 = State
End Function 
'--------------------------------[自旋锁]--------------------------------
Function 自旋锁创建()
    自旋锁创建 = GetThreadID() + Int(999999999 * Rnd)
End Function
Sub 自旋锁进入(锁句柄)
    Dim 锁ID标识, 锁状态标识
    锁ID标识 = "自旋锁_神梦无痕_QQ：1042207232_由【a188652011】提供_线程ID_" & 锁句柄
    锁状态标识 = "自旋锁_神梦无痕_QQ：1042207232_由【a188652011】提供_状态_" & 锁句柄
    Do While GetEnv(锁ID标识) <> GetThreadID()
        While GetEnv(锁状态标识) = True
            Delay Int((600 - 100 + 1) * Rnd + 100)
        Wend
        SetEnv 锁状态标识, True
        SetEnv 锁ID标识, GetThreadID()
        Delay Int((100 - 10 + 1) * Rnd + 10)
    Loop
End Sub
Sub 自旋锁退出(锁句柄)
    Dim 锁ID标识, 锁状态标识
    锁ID标识 = "自旋锁_神梦无痕_QQ：1042207232_由【a188652011】提供_线程ID_" & 锁句柄
    锁状态标识 = "自旋锁_神梦无痕_QQ：1042207232_由【a188652011】提供_状态_" & 锁句柄
    SetEnv 锁ID标识, 0
    SetEnv 锁状态标识, False
End Sub
'--------------------------------[阻塞锁]--------------------------------
Function 阻塞锁创建()
    阻塞锁创建 = GetThreadID() + Int(999999999 * Rnd)
    SetEnv "阻塞锁_神梦无痕_QQ：1042207232_由【a188652011】提供_状态_" & 阻塞锁创建, True
End Function
Sub 阻塞锁进入(锁句柄)
    Dim 锁ID标识, 锁状态标识
    锁ID标识 = "阻塞锁_神梦无痕_QQ：1042207232_由【a188652011】提供_线程ID_" & 锁句柄
    锁状态标识 = "阻塞锁_神梦无痕_QQ：1042207232_由【a188652011】提供_状态_" & 锁句柄
    Do While GetEnv(锁ID标识) <> GetThreadID()
        While GetEnv(锁状态标识) = True  
            Delay Int((500 - 100 + 1) * Rnd + 100)
        Wend
        SetEnv 锁状态标识, True
        SetEnv 锁ID标识, GetThreadID()
        Delay Int((100 - 10 + 1) * Rnd + 10)
    Loop
    SetEnv 锁ID标识, 0
    SetEnv 锁状态标识, False
End Sub
Sub 阻塞锁退出(锁句柄)
    Dim 锁ID标识, 锁状态标识
    锁ID标识 = "阻塞锁_神梦无痕_QQ：1042207232_由【a188652011】提供_线程ID_" & 锁句柄
    锁状态标识 = "阻塞锁_神梦无痕_QQ：1042207232_由【a188652011】提供_状态_" & 锁句柄
    SetEnv 锁ID标识, 0
    SetEnv 锁状态标识, False
End Sub
