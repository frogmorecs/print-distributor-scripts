' Script to switch paper trays
'
' This script will switch paper trays on a Dell 5330 PostScript printer
'
'

Option explicit

'
' The tray codes are
'
' null = Printer auto select
' "1" = Tray 1
' "2" = MP Feeder (MPF)
' "3" = Tray 2
' "4" = Tray 3
' "5" = Tray 4

Dim tray
tray = "3"

Dim fso, input, output
Set fso = CreateObject("Scripting.FileSystemObject")
Set input = fso.OpenTextFile(PrintFilePath, 1)
Set output = fso.CreateTextFile(PrintFilePath & ".tmp", true)

ReplaceFeature "*InputSlot", "<</ManualFeed false /MediaPosition " & tray & ">> setpagedevice", input, output

Sub ReplaceFeature(featureName, postScript, inputStream, outputStream)
	Dim line, foundFeature
	foundFeature = false

	Do While Not inputStream.AtEndOfStream
		line = inputStream.ReadLine
		If IsBeginFeature(line, "*InputSlot") Then
			ReplaceFeatureInStream postScript, inputStream, outputStream
			foundfeature = true
		Else
			If IsEndSetup(line) And Not foundFeature Then
				WriteFeatureInStream postScript, outputStream
			End If
			outputStream.WriteLine line
		End If
	Loop
End Sub

Sub WriteFeatureInStream(postScript, outputStream)
	outputStream.WriteLine "featurebegin{"
	outputStream.WriteLine "%%BeginFeature: *InputSlot"
	outputStream.WriteLine postScript
	outputStream.WriteLine "%%EndFeature"
	outputStream.WriteLine "}featurecleanup"
End Sub

Sub ReplaceFeatureInStream(postScript, inputStream, outputStream)
	outputStream.WriteLine "%%BeginFeature: *InputSlot"
	outputStream.WriteLine postScript
	outputStream.WriteLine "%%EndFeature"
	ReadToEndOfFeature(inputStream)
End Sub

Sub ReadToEndOfFeature(inputStream)
	Dim line
	Do While Not inputStream.AtEndOfStream
		line = inputStream.ReadLine
		If IsEndFeature(line) Then
			Exit Do
		End If
	Loop
End Sub

Function IsBeginFeature(line, feature)
	Dim beginSetup
	beginSetup = "%%BeginFeature: " & feature
	
	IsBeginFeature = InStr(1, line, beginSetup, 1)
End Function

Function IsEndFeature(line)
	Dim endFeature
	endFeature = "%%EndFeature"
	
	IsEndFeature = InStr(1, line, endFeature, 1)
End Function

Function IsEndSetup(line)
	Dim endSetup
	endSetup = "%%EndSetup"

	IsEndSetup = InStr(1, line, endSetup, 1) 
End Function

