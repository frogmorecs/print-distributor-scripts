'Script to print to printer specified in the document name

Option Explicit
Dim tokens, printer

tokens = Split(DocumentName, "-")

If UBound(tokens) >= 2 Then
	printer = "\\" & tokens(1) & "\" & tokens(2)
	LogMessage "Printing to " & printer

	On Error Resume Next
	Reprint PrintFilePath, printer, DocumentName

	If Err.Number <> 0 Then
		LogMessage "Printer not reachable or insufficient permissions"
	End If
Else
	LogMessage "Document name format error: " & DocumentName
End If