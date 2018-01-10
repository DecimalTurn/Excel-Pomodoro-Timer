Attribute VB_Name = "Records"
Option Explicit

Sub Clear_all_records()
    Dim sht As Worksheet
    Set sht = Sheets("Pomodoro")
    
    'Ask the user if an ARCHIVED version should be saved
    Dim Decision As Boolean, ireply As Variant
    ireply = MsgBox(prompt:="Would you like to save your records in an ARCHIVED file.", Buttons:=vbYesNoCancel, Title:="Decision")
    
    If ireply = vbYes Then
        Decision = True
    ElseIf ireply = vbNo Then
        Decision = False
    Else 'They cancelled (VbCancel)
        Exit Sub
    End If
    
    If Decision = True Then
        Call Achive_records(sht)
    End If
    
    'Clear the content of the table
    Dim topleft As Range, bottomright As Range
    Set topleft = sht.Range("A1").End(xlDown).Offset(1, 0)
    Set bottomright = sht.Cells.SpecialCells(xlCellTypeLastCell).Offset(10, 0)
    
    Range(topleft, bottomright).ClearContents
    sht.Range("TaskNameRng").ClearContents

End Sub

Sub new_record_test()

    Call Add_new_record(Date, Now, Now, True, "TaskName")

End Sub

Sub Add_new_record(Pdate, Pstart, Pend, Pcompleted, TaskName)

    Dim sht As Worksheet
    Set sht = Sheets("Pomodoro")
    
    'Find where to put the new line
    Dim rnb As Long
    Dim c As Variant
    For Each c In Range(sht.Cells(Range("TopLeftCorner").Row + 1, 1), sht.Cells(LastCell_row(sht) + 1, 1))
        If IsEmpty(c) Then
            rnb = c.Row
            Exit For
        End If
    Next c
    
    sht.Cells(rnb, 1).Value2 = Pdate
    sht.Cells(rnb, 2).Value2 = Pstart
    sht.Cells(rnb, 3).Value2 = Pend
    sht.Cells(rnb, 4).Value2 = Pcompleted
    sht.Cells(rnb, 5).Value2 = TaskName
    
    'Formatting
    sht.Cells(rnb, 1).NumberFormat = "yyyy-mm-dd"
    sht.Cells(rnb, 2).NumberFormat = "h:mm AM/PM"
    sht.Cells(rnb, 3).NumberFormat = "h:mm AM/PM"
    sht.Cells(rnb, 4).NumberFormat = "General"
    sht.Cells(rnb, 5).NumberFormat = "General"

    Call Add_task(TaskName)
    
End Sub

Sub Add_task(ByVal TaskName As String)

    Dim x As Variant
    On Error Resume Next
    x = Application.Match(TaskName, Range("Recent_Tasks").Value2, 0)
    On Error GoTo 0
    
    If IsError(x) Then
        Sheets("Recent").Cells(LastCell_row(Sheets("Recent")) + 1, 1).Value2 = TaskName
    End If
    
    'Reapply data validation
    On Error Resume Next
    Dim tstrng As Range, tststr As String
    tststr = tstrng.Validation.Formula1
    On Error GoTo 0
    If tststr = vbNullString Then
        With Sheets("Pomodoro").Range("TaskNameRng").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=Recent_Tasks"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    End If
End Sub

Sub Clear_Recent_Tasks()
    
    Range(Sheets("Recent").Cells(2, 1), Sheets("Recent").Cells(LastCell_row(Sheets("Recent")), 1)).ClearContents
    
    'Refill the dummy task names
'    Sheets("Recent").Cells(2, 1).Value2 = "Check emails"
'    Sheets("Recent").Cells(3, 1).Value2 = "Make phone call"
'    Sheets("Recent").Cells(4, 1).Value2 = "Reading"

End Sub

Sub Achive_records(ByRef sht As Worksheet)

    sht.Copy
    With ActiveWorkbook
        .SaveAs FileName:=ThisWorkbook.Path & "\Pomodoro_Timer_ARCHIVE_" & Format(Now, "YYYYMMDD") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=False
    End With

End Sub

Sub Refresh_Summary_PivotTable()
    
    Sheets("Summary").PivotTables("PivotTable1").PivotCache.Refresh

End Sub
