Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
'MsgBox(WScript.Arguments(0) & Vblf & WScript.Arguments(1))
oFso.CopyFile WScript.Arguments(0), WScript.Arguments(1), false