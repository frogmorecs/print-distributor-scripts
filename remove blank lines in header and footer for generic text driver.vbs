' Remove blank lines in header and footer

Dim objFSO, objTSIn, objTSOut

Dim strLine
Dim bInText

bInText = false

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTSIn = objFSO.OpenTextFile(PrintFilePath, 1, false, 0)
Set objTSOut = objFSO.CreateTextFile(PrintFilePath & ".tmp", true, false)

Do While objTSIn.AtEndOfStream = False
	strLine = objTSIn.ReadLine

	strLine = Replace(strLine, Chr(12), Chr(13) & Chr(10))
	strLine = Replace(strLine, Chr(13), "")
	strLine = Replace(strLine, Chr(10), "")


	If bInText Or (Len(Trim(strLine)) > 0) Then
		objTSOut.WriteLine(strLine)
		bInText = true
	End If
Loop

objTSOut.WriteLine("")

objTSIn.Close
objTSOut.Close

objFSO.DeleteFile PrintFilePath
objFSO.MoveFile PrintFilePath & ".tmp", PrintFilePath