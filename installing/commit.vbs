Dim name, fso, dirname, shell

name = "render"

Set fso = CreateObject("Scripting.FileSystemObject")

dirname = fso.BuildPath(fso.GetAbsolutePathName(".."), name)

Set shell = WScript.CreateObject("WScript.Shell")

shell.Run "cmd.exe /k cd " & dirname

Set shell = Nothing
