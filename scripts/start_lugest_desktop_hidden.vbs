Option Explicit

Dim shell, fso, scriptDir, ps1Path, installDir, args, command

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ps1Path = fso.BuildPath(scriptDir, "Arrancar LuisGEST Desktop.ps1")
installDir = ""

If WScript.Arguments.Count > 0 Then
    installDir = WScript.Arguments(0)
End If

args = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1Path & """"
If Len(installDir) > 0 Then
    args = args & " -InstallDir """ & installDir & """"
End If

command = "powershell.exe " & args
shell.Run command, 0, False
