' Script to reprint to a printer based on the users name
Option Explicit

Dim ptr, map

Set map = CreateObject("Scripting.Dictionary")

' Specify the default printer name
ptr = "Printer 1"

' Add the users default printer to this list

map.Add "Tony Edgecombe", "Printer 2"
map.Add "Konstantin Kashutin", "Printer 3"

If map.Exists(UserName) Then
  ptr = map.Item(UserName)
End If

LogMessage "Reprinting document from " & UserName & " To " & ptr
RePrint PrintFilePath, ptr, DocumentName 
