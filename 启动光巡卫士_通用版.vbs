' 光巡卫士 一键启动 · 通用版（首次运行会让你填路径）
' 用法：第一次双击会弹 3 个输入框让你填路径（只需填一次），以后双击直接启动。

Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

ScriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)
ConfigPath = ScriptDir & "\路径配置.ini"

If FSO.FileExists(ConfigPath) Then
    Set cfg = FSO.OpenTextFile(ConfigPath, 1, False, -1)
    content = cfg.ReadAll
    cfg.Close
    PYTHON_PATH = ""
    BACKEND_PATH = ""
    FRONTEND_PATH = ""
    For Each line In Split(content, vbCrLf)
        line = Trim(line)
        If Left(line, 12) = "PYTHON_PATH=" Then
            PYTHON_PATH = Mid(line, 13)
        ElseIf Left(line, 13) = "BACKEND_PATH=" Then
            BACKEND_PATH = Mid(line, 14)
        ElseIf Left(line, 14) = "FRONTEND_PATH=" Then
            FRONTEND_PATH = Mid(line, 15)
        End If
    Next
Else
    PYTHON_PATH = InputBox("请输入 Python 解释器完整路径：", "光巡卫士 · 配置 Python")
    If PYTHON_PATH = "" Then WScript.Quit 0
    BACKEND_PATH = InputBox("请输入后端 server.py 完整路径：", "光巡卫士 · 配置后端")
    If BACKEND_PATH = "" Then WScript.Quit 0
    FRONTEND_PATH = InputBox("请输入前端 index.html 完整路径：", "光巡卫士 · 配置前端")
    If FRONTEND_PATH = "" Then WScript.Quit 0
    Set cfg = FSO.CreateTextFile(ConfigPath, True, False)
    cfg.WriteLine "PYTHON_PATH=" & PYTHON_PATH
    cfg.WriteLine "BACKEND_PATH=" & BACKEND_PATH
    cfg.WriteLine "FRONTEND_PATH=" & FRONTEND_PATH
    cfg.Close
End If

If Not FSO.FileExists(PYTHON_PATH) Then
    MsgBox "找不到 Python：" & vbCrLf & PYTHON_PATH & vbCrLf & vbCrLf & "请删除 路径配置.ini 后再运行。", vbCritical, "启动失败"
    WScript.Quit 1
End If
If Not FSO.FileExists(BACKEND_PATH) Then
    MsgBox "找不到后端文件：" & vbCrLf & BACKEND_PATH, vbCritical, "启动失败"
    WScript.Quit 1
End If
If Not FSO.FileExists(FRONTEND_PATH) Then
    MsgBox "找不到前端文件：" & vbCrLf & FRONTEND_PATH, vbCritical, "启动失败"
    WScript.Quit 1
End If

BackendCmd = """" & PYTHON_PATH & """ """ & BACKEND_PATH & """"
Ret = WshShell.Run(BackendCmd, 0, False)
If Ret <> 0 Then
    MsgBox "启动 Python 后端失败，返回码：" & Ret, vbCritical, "启动失败"
    WScript.Quit 1
End If

WScript.Sleep 3000
WshShell.Run """" & FRONTEND_PATH & """", 1, False

Set WshShell = Nothing
Set FSO = Nothing