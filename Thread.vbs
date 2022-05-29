' Following arguments are used:
' 0 - Workbook name
' 1 - Worksheet name
' 2 - Thread ID
' 3 - Parameter 1 (sleeping time)
' 4 - Output cell
' 5 - State cell

dim i, j, oXL

Set oXL = GetObject(, "Excel.Application")

oXL.workbooks(WScript.Arguments(0)).sheets(WScript.Arguments(1)).Range(WScript.Arguments(4)) = ""
oXL.workbooks(WScript.Arguments(0)).sheets(WScript.Arguments(1)).Range(WScript.Arguments(5)) = "Running"

WScript.Sleep WScript.Arguments(3)*1000

oXL.workbooks(WScript.Arguments(0)).sheets(WScript.Arguments(1)).Range(WScript.Arguments(4)) = "Greetings from thread " & WScript.Arguments(2)
oXL.workbooks(WScript.Arguments(0)).sheets(WScript.Arguments(1)).Range(WScript.Arguments(5)) = "Finished"