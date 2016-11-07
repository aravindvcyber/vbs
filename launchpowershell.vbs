Option Explicit
Dim objFso, objFoo,objFile ,f ,i ,file ,errorCode, filestreamIN  
WScript.Echo "Creating File System Object"
'Create the FileSystemObject and write to file.
Set objFso = CreateObject("Scripting.FileSystemObject")
 
if objFso.FileExists("d:\ps\copy.txt") then
    objFso.DeleteFile "d:\ps\copy.txt", True
    WScript.Echo "Deleting the old files?"
end if 
set objFile = objFso.CreateTextFile("d:\ps\copy.txt", True)
dim objShell
Set objShell = CreateObject("Wscript.shell")
errorCode = objShell.run("powershell -file d:\ps\ps.ps1 -inpUrl http://the-machine/myfdfdee")
If Not errorCode  <> 0 Then
    Wscript.echo "PS successful " & errorCode 
Else
    Wscript.echo "PS Failed " & errorCode 
End if
WScript.Echo "Opening File System Object"
Set filestreamIN = CreateObject("Scripting.FileSystemObject").OpenTextFile("d:\ps\input.txt",1)
file = Split(filestreamIN.ReadAll(), vbCrLf)
'objFile.WriteLine(filestreamIN.ReadAll())
filestreamIN.Close()
Set filestreamIN = Nothing
WScript.Echo "reading from result"
for i = LBound(file) to UBound(file)
    objFile.WriteLine(file(i))
Next
objFile.Close


