Dim name, fso, dirname, shell

name = "render"

Set fso = CreateObject("Scripting.FileSystemObject")

dirname = fso.BuildPath(fso.GetAbsolutePathName(".."), name)

If Not fso.FolderExists(dirname) Then
    fso.CreateFolder(dirname)
End If

Set f = fso.OpenTextFile(dirname & "\README.md", 2, True)
f.Write "### " & name
f.Close

Set f = fso.OpenTextFile(dirname & "\LICENSE", 2, True)
f.WriteLine "MIT License"
f.Write vbCrLf
f.WriteLine "Copyright (c) 2022 askphp"
f.Write vbCrLf
f.WriteLine "Permission is hereby granted, free of charge, to any person obtaining a copy"
f.WriteLine "of this software and associated documentation files (the "" Software ""), to deal"
f.WriteLine "in the Software without restriction, including without limitation the rights"
f.WriteLine "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell"
f.WriteLine "copies of the Software, and to permit persons to whom the Software is"
f.WriteLine "furnished to do so, subject to the following conditions:"
f.Write vbCrLf
f.WriteLine "The above copyright notice and this permission notice shall be included in all"
f.WriteLine "copies or substantial portions of the Software."
f.Write vbCrLf
f.WriteLine "THE SOFTWARE IS PROVIDED "" AS Is "", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR"
f.WriteLine "IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,"
f.WriteLine "FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE"
f.WriteLine "AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER"
f.WriteLine "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,"
f.WriteLine "OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE"
f.WriteLine "SOFTWARE."
f.Close

Set fso = Nothing

Set shell = CreateObject("WScript.Shell")

shell.Run "explorer.exe " & dirname

WScript.Sleep 500

shell.Run "cmd.exe /K ""cd /d " & dirname & """"

Set shell = Nothing
