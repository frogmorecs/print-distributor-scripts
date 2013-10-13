' Log print details to file

Dim oFSO
Dim oTS

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oTS = oFSO.OpenTextFile("C:\Print-log.csv", 8, true)

oTS.WriteLine DocumentName & "," & UserName & "," & TotalPages & "," & Date & "," & LastPrinter

oTS.Close
