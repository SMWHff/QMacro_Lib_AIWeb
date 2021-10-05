Call Main()
Function Main()
    On Error Resume Next
    MsgBox "请先关闭按键精灵和小精灵，然后点击确定开始修复。。。", 1+64+4096, "一键修复"
    Set fso = CreateObject("Scripting.Filesystemobject")
    path = "C:\Windos\System32\SMWH.dll"
    path2 = "C:\Windos\System32\Chrome.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) then
        fso.DeleteFile path
        fso.DeleteFile path2
        If fso.FileExists(path) Or fso.FileExists(path2) then
            MsgBox "修复失败，请先关闭按键精灵和小精灵！", 1 + 64 + 4096, "一键修复"
            Exit Function
        End If
    End If
    path = "C:\Windos\SysWOW64\SMWH.dll"
    path2 = "C:\Windos\SysWOW64\Chrome.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) then
        fso.DeleteFile path
        fso.DeleteFile path2
        If fso.FileExists(path) Or fso.FileExists(path2) then
            MsgBox "修复失败，请先关闭按键精灵和小精灵！", 1 + 64 + 4096, "一键修复"
            Exit Function
        End If
    End If
    path = CreateObject("Shell.Application").Namespace(&H1A).Self.Path & "\MyMacro\plugin\SMWH.dll"
    path2 = CreateObject("Shell.Application").Namespace(&H1A).Self.Path & "\MyMacro\plugin\Chrome.dll"
    path3 = CreateObject("Shell.Application").Namespace(&H1A).Self.Path & "\MyMacro\plugin\SmWeb.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) Or fso.FileExists(path3) then
        fso.DeleteFile path
        fso.DeleteFile path2
        fso.DeleteFile path3
        If fso.FileExists(path) Or fso.FileExists(path2) Or fso.FileExists(path3) then
            MsgBox "修复失败，请先关闭按键精灵和小精灵！", 1 + 64 + 4096, "一键修复"
            Exit Function
        End If
    End If
    path = CreateObject("Shell.Application").Namespace(&H1C).Self.Path & "\Temp\SMWH.dll"
    path2 = CreateObject("Shell.Application").Namespace(&H1C).Self.Path & "\Temp\Chrome.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) then
        fso.DeleteFile path
        fso.DeleteFile path2
        If fso.FileExists(path) Or fso.FileExists(path2) then
            MsgBox "修复失败，请先关闭按键精灵和小精灵！", 1 + 64 + 4096, "一键修复"
            Exit Function
        End If
    End If
    path = "C:\神梦辅助\SMWH\SMWH.dll"
    path2 = "C:\神梦辅助\SMWH\Chrome.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) then
        fso.DeleteFile path
        fso.DeleteFile path2
        If fso.FileExists(path) Or fso.FileExists(path2) then
            MsgBox "修复失败，请先关闭按键精灵和小精灵！", 1 + 64 + 4096, "一键修复"
            Exit Function
        End If
    End If
    Set fso = Nothing
    MsgBox "修复完成！", 1 + 64 + 4096, "一键修复"
End Function