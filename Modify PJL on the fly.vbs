' Script to send documents to Ricoh Aficio SP C820DN Mailboxes

'
' Mail box options are:
'
' UPPER					Standard output bin
' OPTIONALOUTPUTBIN2	Mailbox tray 1
' OPTIONALOUTPUTBIN3	Mailbox tray 2
' OPTIONALOUTPUTBIN4	Mailbox tray 3
' OPTIONALOUTPUTBIN5	Mailbox tray 4
'
'

Option Explicit

PrintToOutputBin "Change this to the printer name", "UPPER"
PrintToOutputBin "Change this to the printer name", "OPTIONALOUTPUTBIN2"
PrintToOutputBin "Change this to the printer name", "OPTIONALOUTPUTBIN3"
PrintToOutputBin "Change this to the printer name", "OPTIONALOUTPUTBIN4"
PrintToOutputBin "Change this to the printer name", "OPTIONALOUTPUTBIN5"

Sub PrintToOutputBin(printer, bin)
	Dim fso, input, output

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set input = fso.OpenTextFile(PrintFilePath, 1, false, 0)
	Set output = fso.CreateTextFile(PrintFilePath & ".tmp", true)

	ProcessToStartOfPJL input, output
	ProcessPJL input, output, bin
	ProcessRemainder input, output

	input.Close
	output.Close

	RePrint PrintFilePath & ".tmp", printer, DocumentName 

	fso.DeleteFile PrintFilePath & ".tmp"
End Sub

Sub ProcessPJL(input, output, bin)
	Dim line
	Dim found

	found = False
	Do While input.AtEndOfStream = False
		line = input.ReadLine

		If InStr(1, line, "@PJL SET OUTBIN") Then
			output.WriteLine "@PJL SET OUTBIN=" & bin
			found = True
		Else
			If InStr(1, line, "PJL ENTER LANGUAGE") Then
				If Not found Then
					output.WriteLine "@PJL SET OUTBIN=" & bin
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