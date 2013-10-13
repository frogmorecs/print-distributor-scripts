' Script to select printer based on paper stock

Option Explicit

Dim PrinterName, strData, pos, posw, PaperType
Dim objStocks, objFSO, objTSIn

' Modify the next line to change the default printer if no stock selected
PrinterName = "Printer 1"

Set objStocks = CreateObject("Scripting.Dictionary")

' Modify the following lines to specify the printer assocaited with each stock
objStocks.Add "Plain", "Printer 1"
objStocks.Add "Preprinted", "Printer 2"
objStocks.Add "Letterhead", "Printer 3"
objStocks.Add "Transparency", "Printer 1"
objStocks.Add "Prepunched", "Printer 2"
objStocks.Add "Labels", "Printer 3"
objStocks.Add "Bond", "Printer 1"
objStocks.Add "Recycled", "Printer 2"
objStocks.Add "Color", "Printer 3"
objStocks.Add "Light 60-75 g/m2", "Printer 1"
objStocks.Add "Cardstock 164-200 g/m2", "Printer 2"
objStocks.Add "Rough", "Printer 3"
objStocks.Add "Envelope", "Printer 1"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTSIn = objFSO.OpenTextFile(PrintFilePath, 1, false, 0)

strData = objTSIn.Read(4096)

pos = 0
Do
	pos = InStr(pos + 1, strData, Chr(27) & "&n")

	If pos > 0 Then
		posw = InStr(pos, strData, "W")

		Size = Mid(strData, pos + 3, posw - pos - 3)

		If IsNumeric(Size) Then
			PaperType = Mid(strData, posw + 2, Size - 1)
			If objStocks.Exists(PaperType) Then
				PrinterName = objStocks.Item(PaperType)
			End If
			Exit Do
		End If
	End If

Loop While pos > 0

objTSIn.Close

RePrint PrintFilePath, PrinterName, DocumentName