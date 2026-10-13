Option Explicit

Dim fso, shell, projectDir, pythonw, scriptPath, command
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = fso.BuildPath(projectDir, ".venv\Scripts\pythonw.exe")
scriptPath = fso.BuildPath(projectDir, "sheet_watcher.py")

If Not fso.FileExists(pythonw) Then
    MsgBox "找不到虚拟环境的 pythonw.exe：" & vbCrLf & pythonw, vbExclamation, "Google Sheet Watcher"
    WScript.Quit 1
End If

shell.CurrentDirectory = projectDir
command = """" & pythonw & """ """ & scriptPath & """ --all --owned-only --quiet"
shell.Run command, 0, False
