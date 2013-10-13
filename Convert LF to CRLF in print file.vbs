' Convert LF to CRLF in print file

Dim objFSO, objTSIn, objTSOut

Dim char


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTSIn = objFSO.OpenTextFile(PrintFilePath, 1, false, 0)
Set objTSOut = objFSO.CreateTextFile(PrintFilePath & ".tmp", true, false)

Do While objTSIn.AtEndOfStream = False
	char = objTSIn.Read(1)

	If Asc(char) = 10 Then
		objTSOut.WriteLine("")
	Else
		objTSOut.Write(char)
	End If
Loop

objTSOut.WriteLine("")

objTSIn.Close
objTSOut.Close

objFSO.DeleteFile PrintFilePath
objFSO.MoveFile PrintFilePath & ".tmp", PrintFilePath