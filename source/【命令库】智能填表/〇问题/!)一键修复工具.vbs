Call Main()
Function Main()
    On Error Resume Next
    MsgBox "���ȹرհ��������С���飬Ȼ����ȷ����ʼ�޸�������", 1+64+4096, "һ���޸�"
    Set fso = CreateObject("Scripting.Filesystemobject")
    path = "C:\Windos\System32\SMWH.dll"
    path2 = "C:\Windos\System32\Chrome.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) then
        fso.DeleteFile path
        fso.DeleteFile path2
        If fso.FileExists(path) Or fso.FileExists(path2) then
            MsgBox "�޸�ʧ�ܣ����ȹرհ��������С���飡", 1 + 64 + 4096, "һ���޸�"
            Exit Function
        End If
    End If
    path = "C:\Windos\SysWOW64\SMWH.dll"
    path2 = "C:\Windos\SysWOW64\Chrome.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) then
        fso.DeleteFile path
        fso.DeleteFile path2
        If fso.FileExists(path) Or fso.FileExists(path2) then
            MsgBox "�޸�ʧ�ܣ����ȹرհ��������С���飡", 1 + 64 + 4096, "һ���޸�"
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
            MsgBox "�޸�ʧ�ܣ����ȹرհ��������С���飡", 1 + 64 + 4096, "һ���޸�"
            Exit Function
        End If
    End If
    path = CreateObject("Shell.Application").Namespace(&H1C).Self.Path & "\Temp\SMWH.dll"
    path2 = CreateObject("Shell.Application").Namespace(&H1C).Self.Path & "\Temp\Chrome.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) then
        fso.DeleteFile path
        fso.DeleteFile path2
        If fso.FileExists(path) Or fso.FileExists(path2) then
            MsgBox "�޸�ʧ�ܣ����ȹرհ��������С���飡", 1 + 64 + 4096, "һ���޸�"
            Exit Function
        End If
    End If
    path = "C:\���θ���\SMWH\SMWH.dll"
    path2 = "C:\���θ���\SMWH\Chrome.dll"
    If fso.FileExists(path) Or fso.FileExists(path2) then
        fso.DeleteFile path
        fso.DeleteFile path2
        If fso.FileExists(path) Or fso.FileExists(path2) then
            MsgBox "�޸�ʧ�ܣ����ȹرհ��������С���飡", 1 + 64 + 4096, "һ���޸�"
            Exit Function
        End If
    End If
    Set fso = Nothing
    MsgBox "�޸���ɣ�", 1 + 64 + 4096, "һ���޸�"
End Function