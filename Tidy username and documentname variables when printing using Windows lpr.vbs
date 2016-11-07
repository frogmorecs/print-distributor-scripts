' Remove IP address from user name on LPD jobs

Option Explicit

Dim Pos
Pos = InStr(1, UserName, "(")

If Pos > 0 Then
	UserName = Trim(Left(UserName, Pos - 1))
End If

' Remove directory information from DocumentName when it is a path

Dim Segments
Segments = Split(DocumentName, "\")

DocumentName = Segments(UBound(Segments))
