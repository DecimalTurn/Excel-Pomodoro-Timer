Attribute VB_Name = "Records"
Option Explicit

Public Sub Clear_all_records()
'PURPOSE: Copy current records to Archive sheet (if user selects yes) and clear records

    Dim sht As Worksheet, arcsht As Worksheet
    Set sht = ThisWorkbook.Sheets("Pomodoro")
    Set arcsht = ThisWorkbook.Sheets("Archive")
    
    'Ask the user if an ARCHIVED version should be saved
    Dim Decision As Boolean, ireply As Variant
    ireply = MsgBox(prompt:="Would you like to save your records to the Archive sheet.", Buttons:=vbYesNoCancel, Title:="Decision")
    
    If ireply = vbYes Then
        Decision = True
    ElseIf ireply = vbNo Then
        Decision = False
    Else 'They cancelled (VbCancel)
        Exit Sub
    End If
       
    Dim topleft As Range, bottomright As Range, SrcRange As Range
    Set topleft = sht.Range("A1").End(xlDown).Offset(1, 0)
    Set bottomright = sht.Cells.SpecialCells(xlCellTypeLastCell).Offset(10, 0)
    Set SrcRange = sht.Range(topleft, bottomright)
       
    If Decision = True Then

        'We need to remove error filtering to avoid errors when inserting data
        On Error Resume Next
          arcsht.ShowAllData
        On Error GoTo 0

        'Destination Range (Just the top left cell)
        'Find where to add the lines for archive
        Dim rnb As Long
        Dim c As Variant
        For Each c In arcsht.Range(arcsht.Cells(arcsht.Range("TopLeftCornerA").Row + 1, 1), arcsht.Cells(LastCell_row(arcsht) + 1, 1))
            If IsEmpty(c) Then
                rnb = c.Row
                Exit For
            End If
        Next c
        
        Dim DestTopLeftRange As Range
        Set DestTopLeftRange = arcsht.Cells(rnb, 1)
        
        VBACopyPaste SrcRange, DestTopLeftRange

    End If
        
    'Clear the content of the table
    SrcRange.ClearContents

End Sub

Public Sub new_record_test()

    Add_new_record Date, Now, Now, True, "TaskName"

End Sub

Public Sub Add_new_record(Pdate As Date, Pstart As Date, Pend As Date, Pcompleted As Boolean, TaskName As String)

    Dim sht As Worksheet
    Set sht = ThisWorkbook.Sheets("Pomodoro")
    
    'Find where to put the new line
    Dim rnb As Long
    Dim c As Variant
    For Each c In sht.Range(sht.Cells(sht.Range("TopLeftCorner").Row + 1, 1), sht.Cells(LastCell_row(sht) + 1, 1))
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

    Add_task TaskName
    
End Sub

Public Sub Add_task(ByVal TaskName As String)

    Dim recSht As Worksheet
    Set recSht = ThisWorkbook.Sheets("Recent")

    Dim x As Variant
    On Error Resume Next
    x = Application.Match(TaskName, recSht.Range("Recent_Tasks").Value2, 0)
    On Error GoTo 0
    
    If IsError(x) Then
        recSht.Cells(LastCell_row(recSht) + 1, 1).Value2 = TaskName
    End If
    
End Sub

Public Sub Clear_Recent_Tasks()
    
    Dim recSht As Worksheet
    Set recSht = ThisWorkbook.Sheets("Recent")
    recSht.Range(recSht.Cells(2, 1), recSht.Cells(LastCell_row(recSht), 1)).ClearContents
    
    'Refill the dummy task names
'    Sheets("Recent").Cells(2, 1).Value2 = "Check emails"
'    Sheets("Recent").Cells(3, 1).Value2 = "Make phone call"
'    Sheets("Recent").Cells(4, 1).Value2 = "Reading"

End Sub

Public Sub Export_records()

    ThisWorkbook.Sheets("Archive").Copy
    With ActiveWorkbook
        .ActiveSheet.Buttons("Button 1").Delete
        .SaveAs FileName:=ThisWorkbook.Path & "\Pomodoro_Timer_ARCHIVE_" & Format$(Now, "YYYYMMDD") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=False
    End With

End Sub

Public Sub Refresh_Summary_PivotTable()
    
    ThisWorkbook.Sheets("Summary").PivotTables("PivotTable1").PivotCache.Refresh

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VBA Utilities
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub VBACopyPaste(ByRef SrcRange As Range, ByRef DestTopLeftRange As Range)
    
    Dim DestRange As Range
    Dim VBAClipBoard() As Variant

    Set DestRange = DestTopLeftRange.Cells(1, 1) 'Must be one cell
    VBAClipBoard = SrcRange
    Dim destSht As Worksheet
    Set destSht = DestTopLeftRange.Parent
    destSht.Range(DestTopLeftRange, DestTopLeftRange.Offset(UBound(VBAClipBoard, 1) - 1, UBound(VBAClipBoard, 2) - 1)).Value2 = VBAClipBoard
    Erase VBAClipBoard

End Sub
