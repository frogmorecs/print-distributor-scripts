Option Explicit

Dim regex, matches, match

Set regex = new RegExp

regex.Pattern = "\d{2}.\d{2}.\d{4}"
Set matches = regex.Execute(DocumentName)

For Each match in matches
	NotifyName = match.Value
	Exit For
Next
