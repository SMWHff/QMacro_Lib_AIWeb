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
MacroID=6cacb99f-5475-404c-a828-669112a686b4
Description=例子1：百度一下
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
UEsDBBQAAgAIAOgcNFPegsu56AQAAAg0AAAJABEAVUlQYWNrYWdlVVQNAAeIAkhhiAJIYYgCSGHtWktvU0cUPtdOiO2YvJo4aQJNCA0BA6HhXV5SjBo1UkoRdkFCKpIJpqA6DrJdCBX/gzXqskJQqd0gFlkgNqAWqYtuW6n9AWxAXRDzzX3gSZQoc2YG36bKiU7GsT33O/M4j28mUfLkxW+df977qf8vWianKEqLtThtkt5zfHWlgyji/71Yq9WCt2sbsq7kDbTJX0Ox1s3QFmgcGoMmoJuhrdAktM1bemqHdkI/gHZBu6EpaA+0139en99uyH9XztEcfqo0RJ9RCW2ZbhNHUtgxwbOcNb778O7L5OD9P5woXi9s9t77iqboE9KXGDlOgN+7Bm7Qyp+dphxNG1iQQAQM8Js18HOY7xtUoDOUp1m0XOmiiDv+Tj/2qvZr8lvhn/9X3ZC1ZXKuPGvgfiT6jxv0d8hMTPubys8Pn/9jMn6Gy76X/rpxpx7/IktqPtU1C2KlyDp5ukxFTRvaEP9a/TpFFT8ixb/zdJ0qUF0LOjTGH5XwzwE1T1foS8xDkZl7g/if9J+pit8k4eeAP4+8b7D+Dnf8zdL6n0b1UYSWNS3p1sDfJI2/jp/BOszQt3x8t15uZ8x/y7L1nzGcf1FKxRj4MfJqeFm+hlO8ovUnLSHjv1p5zt209EV2KJsvVYayhfL1q+8Jf+THS0b9LyLylBF/CvitIyL+9vg+pbr/4tL+n4bv8b1u6f7v8J+nip9Y0f+zsOMa5oGXDfvg/wEHVsVvlfAnUP3fcHNAGVqCDbwslML44+TxclX8pIQ/ibGXzOIPO/+JeHXi3fiLyL7fwAYx7yWd+M/Gb1uy/nnMfxU2iHnIuOsQZIL63lhddmL+U1Q/91LBb5fwM+7zPf9TwVth/zndTPwOCd80fpjWn6Yynb9cKJr2Nyjg1zt/yWHnV93aV2f3ufGH7X+d0v47C/x5oBdRhetImwZ+1zJ8UX99h3HrnML0aeDL5+Xe+ZP+BkyQ4wg+swvav8p3GnH+tIPU84/Il0H9Hwb/FPEyuNewwT+H0A4yxt9DVvmns41EHlLHT5Fd/jng+oE6fi+Fyz+FrbF3/td4/vkh2eWfW9B+zJj/fgn/7wcLC09f//BId/7Drh9GocL/LkD3avQPO/9GQ8a3cf4wjHYrY/8NUP38wQL/ZMf/LWSXf474Pq2Kv5Ws8k83/owy8D8iu/xTPG87A3+QwuWfIl6EyT+3kZz/vHmvYBXybg3Ks0Gn/hyW8MM+vzIV0/idKc/dqhQyc/Om/cfDsd+0/wSYTxU+mIUXfK9X/7P33/Yl8U/wn/0G/l+/f9+zyncawX844xe1Wtx/bYH/uPhHST3+jpB3BirEBv8RqzfOwN9B4d6/jZJd/nMAbZox/p1kn/8cY+Avhn1oJ4mpKefJq+ciEdKKIpGQxx92/W+L/55i7L9dZJf/jqE9zMBPk93714No9zHwd5NV/uN8Sh73VcXfQ1b5Dx3xx6SKv5fs3r+JWv44A3+M7PIf8d+Dhxj4+8gq/2HXH8LegP+Y1v/P/v39xa9PHi/eSY+lw4hfpudP6V+uGfUfuJcL1X5TOYtdXyX9OUgY3n9MovYrwoJGnv+IWjXpv84C+7Z29Suew8ffL/n/GTf2zWIVitr3v1z8AxL+58C7qX377MW/k8S7fz64LP7NgINWGnj/eEjCnwK6iMB66HrnP4clfBvCxZflLVBLAQIXCxQAAgAIAOgcNFPegsu56AQAAAg0AAAJAAkAAAAAAAAAAAAAgAAAAABVSVBhY2thZ2VVVAUAB4gCSGFQSwUGAAAAAAEAAQBAAAAAIAUAAAAA


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


授权码 = "0123456789ABCDEF" ' 填入你自己的授权码

//记得要先初始化命令库，才能用命令库中的功能
If LibE.神梦_智能填表._初始化(授权码, "") = False  Then 
	MsgBox "出错，智能填表初始化失败！", 1 + 16 + 4096, "警告！"
	ExitScript
End If


//启动指定浏览器
Path = Form1.BrowseBox1.Path
If LibE.神梦_智能填表.绑定(Path) Then 
	
	'打开网址导航
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
