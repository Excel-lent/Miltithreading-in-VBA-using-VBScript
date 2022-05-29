Attribute VB_Name = "MainModule"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
#End If

Dim maxThreads%, nTasks%, wsThreads As Worksheet
Dim threads() As New thread

Sub RunAllTasksAsynchronously()
    Dim thread%, execTime_s!, lastUsedRow&
    
    Set wsThreads = ThisWorkbook.Worksheets("Threads")
    
    maxThreads = wsThreads.Cells(1, 5).Value2
    nTasks = wsThreads.Cells(1, 2).Value2
    ReDim threads(maxThreads)
    
    '---Clear previous values---
    lastUsedRow = wsThreads.Range("B" & wsThreads.Rows.Count).End(xlUp).Row
    If lastUsedRow > 2 Then wsThreads.Range("B3:B" & lastUsedRow).ClearContents
    
    For thread = 1 To Application.WorksheetFunction.Min(maxThreads, nTasks)
        ' set execution time of the 1st thread to 1 second, 2nd to 2 seconds, 3rd to 4 seconds and 4th to 8 seconds.
        execTime_s = 2 ^ (thread - 1)
        Call threads(thread).Constructor(wsThreads, thread, execTime_s, 5)    ' 5 is a thread state column.
        Call threads(thread).StartVBScriptThread
    Next thread
    
    MainLoop
End Sub

' Wait until all threads are finished
Sub MainLoop()
    Dim i&, startedTasks&, finishedTasks&, allTasks_Completed As Boolean
    
    startedTasks = maxThreads
    finishedTasks = 0
    allTasks_Completed = False
    
    While Not allTasks_Completed
        DoEvents
        
        For i = 1 To maxThreads
            If threads(i).GetThreadState = "Finished" Then
                finishedTasks = finishedTasks + 1
                
                ' Copy output to the corresponding cell.
                wsThreads.Cells(finishedTasks + 2, 2).Value2 = wsThreads.Cells(i + 2, 6).Value2
                
                ' Start new task
                If startedTasks < nTasks Then
                    startedTasks = startedTasks + 1
                    Call threads(i).StartVBScriptThread
                Else
                    ' If the maximal number of tasks is reached, then set "Thread states" to "" to avoid additional outputs.
                    wsThreads.Cells(i + 2, 5).Value2 = ""
                End If
                
                If startedTasks = finishedTasks Then
                    allTasks_Completed = True
                End If
            End If
        Next i
        
        DoEvents
        Sleep (100)
    Wend
End Sub
