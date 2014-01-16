' Script to print a set number of copies across a group of printers

Dim printers, count
Set printers = CreateObject("Scripting.Dictionary")

' Add an entry here for each printer
addPrinter "Printer 1"
addPrinter "Printer 2"
addPrinter "Printer 3"

' Set the number of rperints here
count = 100

for i = 0 to count - 1
	RePrint PrintFilePath, printers(i Mod printers.Count), DocumentName
next

function AddPrinter(name)
    printers.add printers.count, name
end function