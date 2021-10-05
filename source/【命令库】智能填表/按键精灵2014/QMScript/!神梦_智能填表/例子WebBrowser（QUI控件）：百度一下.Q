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
MacroID=332b8bd7-7414-45e7-a880-5b54155670f5
Description=例子WebBrowser（QUI控件）：百度一下
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
UEsDBBQAAgAIAPxyM1Ol5ELKawMAAAgkAAAJABEAVUlQYWNrYWdlVVQNAAfiR0dh4kdHYeJHR2HtWEtPU0EU/m7Low8s5VWwPKwP1CXujdESSUzwEVo0MdGkYFViKeSCEf03rI2JJv4ME01cuHGhC/es2Ek9c+9cOyUQ58wMNk36kcOUS+d+c87MeU0cIb58Hvqx+yH/E4dwDXEcNJLoU555UgJkgZj8+6DRaESPG110FH6T9Mg9FHvdS9JPkiRJkKRITpGkSQZIMuHWY1DOOeiasKOxhA362UYBN1Gn0cdrcJCjExO9y/vHd6/vfv0+8+6bFxefE+GzZdzCHMyRgOdF/LFjvjP3JuSN+NX/zaOMRYsVpCT/uPQdLn+Z7L2JKu6ggnUauRhGLOAfkrFXd16PHCMfbqd00T4sbPjrFu4HMf+KxXwPdrCdb4uP7z/9stGf4bInMt807jTjX6yl5tPdsyhWiqxTwQpqhmvIUPxLyzpFlz+mxL/7WMMWiekKsgb6xxX+JWKt4Anukh1qzNwbxf8B+U5d/h6Fv0z8O5T3Lfbf4+rfq+z/PFUfNRLfcCWjBvx9iv5N/iLtwype8PmDenmQYf/+Q/u/aml/UZ8nGPwJhDW8ikfkFPvoPPS3mX//aJsHael2qVCq1LcKpaq/9vSE+GffPraa/5Aij0/xp0q/TSDi75j0Kd3zl1TO/yL5Ht/rWs9/Vr5Plz91pP+XaB3PyQ68bDhB/j8CEYf1+dMK/w2q/jeDHOCT1GkNvCyUI/2TCPtyXf4BhX+BdK/bxR92/hPx6upf/WuUfZ/RGoTd6ybxn82fadn/Ctl/m9Yg7FAM9iHKBM2zcTwuk/1zaN576fAPKvzF4P2h/+nwHXH+vFEmf1bht40ftvWnLR5UV4r+xqutqm8736yK7/T+pUwnfzuofU1OXxB/2P43pJy/e8S/Q+w1qsJNkDHgHz7EL+qvl6S3yS3MhAG/el8e3j+Zd5ApOkERd/6Y70T3Tid5/8TRfwRhDhZw0H+y+UW8TMvPLvrPAo1noJ9/x+C0//TOghf/c3Daf2Iy8AN9/cfR3v6z1f/+f/85Aaf9pzdF4wWG/U8r/LbY29tz8h7T/rPd+TPeZn7b/tPF/cM5GqcZ5y+P5v2Dg/6TnX8n4bT/xKz0KV3+KTjtP72L4MX/aTjtPzFD43mG/jNw2n96l5j6i1wd9Z8O+i/2+Sso+i+TBy5a+H/Csv607f+66Gz8AVBLAQIXCxQAAgAIAPxyM1Ol5ELKawMAAAgkAAAJAAkAAAAAAAAAAAAAgAAAAABVSVBhY2thZ2VVVAUAB+JHR2FQSwUGAAAAAAEAAQBAAAAAowMAAAAA


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


//绑定 WebBrowser QUI浏览器控件【首先 在界面上放1个“浏览器控件”】
If LibE.神梦_智能填表.绑定("Web.ocx") Then 

    '打开百度
    Call LibE.神梦_智能填表.网页_跳转("https://www.baidu.com/")
    Delay 2000  '强制等待 2 秒
    
    '输入“神梦填表”
    Call LibE.神梦_智能填表.元素_置输入框("神梦填表", "id:kw") 
    Delay 2000  '强制等待 2 秒
    
    '点击“百度一下”按钮
    Call LibE.神梦_智能填表.元素_点击("id:su")
    Delay 2000  '强制等待 2 秒
    
    MsgBox "脚本执行完毕！", 1+64+4096, "百度一下"
Else 
    MsgBox "出错，绑定浏览器失败！" & vbCrLf & vbCrLf &"失败原因：" & vbCrLf & [错误提示], 1 + 16 + 4096, "警告！"
    ExitScript 
End If
