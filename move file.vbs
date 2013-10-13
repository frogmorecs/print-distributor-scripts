' Script to move print file

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile PrintFilePath, "C:\Archive\" & SerialNumber & ".prn"
