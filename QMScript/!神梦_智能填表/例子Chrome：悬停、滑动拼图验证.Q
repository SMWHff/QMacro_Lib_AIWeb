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
MacroID=30521425-5e98-403b-b4ba-5cf3cbb80120
Description=例子Chrome：悬停、滑动拼图验证
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
    
    '打开网址
    Call LibE.神梦_智能填表.网页_跳转("https://www.imusic.cn/#/ring/openring")
    Call LibE.神梦_智能填表.网页_是否就绪(30)
    Delay 2000  '强制等待 2 秒
    
    
    //鼠标悬停，显示拼图
    Call LibE.神梦_智能填表.元素_悬停("div.verify_block")
    Delay 2000  '强制等待 2 秒
    
    //通过技术手段获取起点和终点坐标，可以对接打码平台获取
    起点坐标X = 260 : 起点坐标Y = 420
    终点坐标X = 530 : 终点坐标Y = 428
    
    //带轨迹滑动
    Call LibE.神梦_智能填表.元素_鼠标按下(起点坐标X, 起点坐标Y, "div.verify_block")
    Delay 100
    For i = 起点坐标X To 终点坐标X Step ep + 10
        ep = ep + 1
        yc = yc + 1
        Call LibE.神梦_智能填表.元素_鼠标移动(i, 起点坐标Y, "div.verify_block")
        'MoveTo i, 起点坐标Y
        Delay yc + 10 * yc
    Next
    Call LibE.神梦_智能填表.元素_鼠标弹起(终点坐标X, 终点坐标Y, "div.verify_block")
    Delay 3500  '强制等待 3.5 秒
    
    MsgBox "脚本执行完毕！", 1+64+4096, "填表例子"
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If
