' Add blank page to end of documents with odd number of pages

Option Explicit

Dim fso, ts

If TotalPages mod 2 = 1 Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(PrintFilePath, 8)

	ts.WriteLine "%!PS-Adobe-3.0"
	ts.WriteLine "showpage"

	ts.Close
End If
