Option Explicit

Dim timeout
' Set the timeout value in seconds
'

timeout = 60


AbortJob = AbortDocument()
SaveDocumentName()

If AbortJob Then
	LogMessage "Skipping duplicate document: " & DocumentName
End If

Sub SaveDocumentName()
    Dim fso, ts
    Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.CreateTextFile(DNPath(), true, true)
	ts.WriteLine(DocumentName)
	ts.Close
End Sub

Function AbortDocument()
    Dim fso, pf, lf
    Set fso = CreateObject("Scripting.FileSystemObject")
    If not fso.FileExists(DNPath()) Then
        AbortDocument = false
        Exit Function
    End If

    Set pf = fso.GetFile(PrintFilePath)
    Set lf = fso.GetFile(DNPath())

    If Abs(DateDiff("s", pf.DateLastModified, lf.DateLastModified)) > timeout Then
        AbortDocument = false
        Exit Function
    End If

    AbortDocument = (DocumentName = LastDocumentName())
End Function

Function LastDocumentName()
    Dim fso, ts
    Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(DNPath(), 1, false, -1)
    LastDocumentName = ts.ReadLine()
	ts.Close
End Function

Function DNPath()
    DNPath = RawFolder & "lastdocumentname.tmp"
End Function

Function RawFolder()
    Dim pos

    pos = InStrRev(PrintFilePath, "\")
    If (pos > 0) Then
        RawFolder = Left(PrintFilePath, pos)
    Else
        RawFolder = "C:\"
    End If
End Function