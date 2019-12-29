Dim FSO, vbStr, Arr, Str, Counter, vbVar
vbVar = ""
Set FSO = CreateObject("Scripting.FileSystemObject")
vbStr = FSO.OpenTextFile("reply.ini").ReadAll
Arr = Split(vbStr, vbCrLf & vbCrLf)
For Each Str In Arr
  If InStr(vbVar, Str) = 0 Then vbVar = vbVar & Str & vbCrLf & vbCrLf   
Next
FSO.OpenTextFile("reply.txt", 2, True).Write vbVar
CreateObject("Wscript.Shell").run "cmd /cstart reply.txt", True, False
Set FSO = Nothing