' Script to set user name in PJL

Option Explicit

Main()

Sub Main()
	Dim fso, input, output

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set input = fso.OpenTextFile(PrintFilePath, 1, false, 0)
	Set output = fso.CreateTextFile(PrintFilePath & ".tmp", true)

	ProcessToStartOfPJL input, output
	ProcessPJL input, output
	ProcessRemainder input, output

	input.Close
	output.Close

	fso.CopyFile PrintFilePath & ".tmp", PrintFilePath, True
	fso.DeleteFile PrintFilePath & ".tmp"
End Sub



Sub ProcessPJL(input, output)
	Dim line
	Dim found

	found = False

	Do While input.AtEndOfStream = False
		line = input.ReadLine

		If InStr(1, line, "PJL COMMENT user") Then
			output.WriteLine line
			found = True
		Else
			If InStr(1, line, "PJL ENTER LANGUAGE") Then
				If Not found Then
					output.WriteLine "@PJL COMMENT user=" & UserName
				End If
				output.WriteLine line
				Exit Sub
			End If
			output.WriteLine line
		End If
	Loop
End Sub

Sub ProcessRemainder(input, output)
	Do While input.AtEndOfStream = False
		output.Write input.Read(1)
	Loop
End Sub

Sub ProcessToStartOfPJL(input, output)
	Dim strPJLStartSequence, i, c
	strPJLStartSequence = Chr(27) & "%-12345X"

	Do While input.AtEndOfStream = False
		For i = 1 To Len(strPJLStartSequence)
			c = input.Read(1)
			output.Write(c)
			If c <> Mid(strPJLStartSequence, i, 1) Then
				Exit For
			End If

			'Have we read the whole PJL exit language sequence?
			If i = Len(strPJLStartSequence) Then
				Exit Sub
			End If
		Next
	Loop
End Sub