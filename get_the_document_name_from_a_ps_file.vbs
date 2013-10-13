' Script to extract the document name from a PostScript file

Option Explicit

Dim objFSO ' Scripting.FileSystemObject
Dim objTS ' Scripting.TextStream
Dim strLine
Dim i

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTS = objFSO.OpenTextFile(PrintFilePath, 1, false, 0)

For i = 1 to 100
	If objTS.AtEndOfStream Then
		Exit For
	End If

	strLine = Trim(objTS.ReadLine)

	If Left(strLine, 8) = "%%Title:" Then
		DocumentName = Trim(Mid(strLine, 9))
		Exit For
	End If
Next

If Left(DocumentName, 1) = "(" Then
	DocumentName = Mid(DocumentName, 2, Len(DocumentName) - 2)
End If
