scriptPath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & scriptPath & "\pulseaudio.bat" & Chr(34), 0
Set WshShell = Nothing