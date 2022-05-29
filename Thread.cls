VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim sFileName$, threadId%, param1!, threadStateColumn%, worksheetName$, usedWorksheet As Worksheet

Public Sub Constructor(worksheet_ As Worksheet, threadId_%, executionTime_s!, threadStateColumn_%)
    Dim wsh As Object
    sFileName = ActiveWorkbook.Path & "\Thread.vbs"
    threadId = threadId_
    param1 = executionTime_s
    threadStateColumn = threadStateColumn_
    Set usedWorksheet = worksheet_
    worksheetName = usedWorksheet.Name
End Sub

' Started thread can be controlled only via commonly used cells.
Public Sub StartVBScriptThread()
    Dim wsh As Object
    
    usedWorksheet.Cells(threadId + 2, threadStateColumn).Value2 = ""
    Set wsh = VBA.CreateObject("WScript.Shell")
    ' Run VBscript in background with following parameters:
    ' Workbook name
    ' Worksheet name
    ' ThreadID
    ' Input parameter nr. 1
    ' Output cell
    ' Cell with thread state
    wsh.Run """" & sFileName & " """"" & ThisWorkbook.Name & """ " & usedWorksheet.Name & " " & threadId & _
            " " & param1 & " F" & CStr(threadId + 2) & " E" & CStr(threadId + 2) & """"
    
    Set wsh = Nothing
End Sub

Function GetThreadState() As String
    GetThreadState = usedWorksheet.Cells(threadId + 2, threadStateColumn).Value2
End Function
