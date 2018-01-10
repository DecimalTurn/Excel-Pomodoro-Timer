Attribute VB_Name = "Assembler"
Option Explicit
'Instructions

'1. Create an Excel file called Assembler.xlsm (for example) in the same folder as Installer.bas:
'   *\Excel-Pomodoro-Timer\

'2. Open the VB Editor (Alt+F11) right click on the Installer VB Project and choose Import a file and chose:
'    *\Excel-Pomodoro-Timer\Assembler.bas

'3. Run Assemble from the module Assembler (Click somewhere inside the macro and press F5).
'   Make sure to wait for the confirmation message at the end before doing anything with Excel.

'4. Use the tool vbaDeveloper (Available here: https://github.com/DecimalTurn/vbaDeveloper/releases) to import the VBA code.
'   - Open vbaDeveloper.xlam
'   - Look at the Add-ins ribbon and choose: vbaDeveloper > Import code for ... > Pomodoro_Timer.xlsb

'5. Save the file

Public Const SHORT_NAME = "Pomodoro_Timer"
Public Const EXT = ".xlsb"

Sub Assemble()

    If testFileLocation = False Then
        Exit Sub
    End If

    On Error Resume Next
    Workbooks(SHORT_NAME & EXT).Close
    On Error GoTo 0
    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    Dim sht As Worksheet
    Dim sht1 As Worksheet
    Dim sht2 As Worksheet
    Dim sht3 As Worksheet
    Dim sht4 As Worksheet
    
    Set sht1 = wb.Sheets.Add
    sht1.Name = "Pomodoro"
    'Delete any other sheet
    For Each sht In wb.Sheets
        If sht.Name <> sht1.Name Then
            Application.DisplayAlerts = False
            sht.Delete
            Application.DisplayAlerts = True
        End If
    Next
    'Create the remaining sheets
    Set sht2 = wb.Sheets.Add(After:=Sheets(Sheets.Count))
    sht2.Name = "Summary"
    Set sht3 = wb.Sheets.Add(After:=Sheets(Sheets.Count))
    sht3.Name = "Recent"
    Set sht4 = wb.Sheets.Add(After:=Sheets(Sheets.Count))
    sht4.Name = "Settings"
    
    Range("A1").Select
    
    '*******************************
    'Sheet Pomodoro
    '*******************************
    Set sht = sht1
    sht.Select
    
    'Column width
    sht.Columns(1).ColumnWidth = 61.2 / 5
    sht.Columns(2).ColumnWidth = 83.4 / 5
    sht.Columns(3).ColumnWidth = 83.4 / 5
    sht.Columns(4).ColumnWidth = 70.8 / 5
    sht.Columns(5).ColumnWidth = 215.4 / 5
    sht.Columns(6).ColumnWidth = 150.6 / 5
       
    'Cells values
    sht.Cells(2, 4).Value2 = "Task Name:"
    sht.Cells(8, 1).Value2 = "Date"
    sht.Cells(8, 2).Value2 = "Start"
    sht.Cells(8, 3).Value2 = "End"
    sht.Cells(8, 4).Value2 = "Completed"
    sht.Cells(8, 5).Value2 = "Task"
    sht.Cells(8, 6).Value2 = "Comment"
    
    sht.Cells(9, 1).Value2 = DateSerial(2017, 11, 25)
    sht.Cells(10, 1).Value2 = DateSerial(2017, 11, 25)
    sht.Cells(11, 1).Value2 = DateSerial(2017, 11, 25)
    
    sht.Cells(9, 2).Value2 = 43064.60846
    sht.Cells(10, 2).Value2 = 43064.69906
    sht.Cells(11, 2).Value2 = 43064.72807

    sht.Cells(9, 3).Value2 = 43064.62582
    sht.Cells(10, 3).Value2 = 43064.71642
    sht.Cells(11, 3).Value2 = 43064.74543
    
    sht.Cells(9, 4).Value2 = True
    sht.Cells(10, 4).Value2 = True
    sht.Cells(11, 4).Value2 = True

    sht.Cells(9, 5).Value2 = "Check emails"
    sht.Cells(10, 5).Value2 = "Make phone call"
    sht.Cells(11, 5).Value2 = "Reading"
    
    'Bulk formatting
    sht.Range(Cells(9, 1), Cells(9, 1).End(xlDown).End(xlDown)).NumberFormat = "YYYY-MM-DD"
    sht.Range(Cells(9, 2), Cells(9, 3).End(xlDown).End(xlDown)).NumberFormat = "HH:MM AM/PM"

    'Cell borders
    Dim x As Variant
    With sht.Cells(2, 5)
        For Each x In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With .Borders(x)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Next x
    End With
    
    'Data validation
    Range("E2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Recent!$A$2:$A$10"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    'Table
    sht.ListObjects.Add(xlSrcRange, sht.Range("$A$8:$F$200"), , xlYes).Name = "Table24"
    
    'Buttons
    Dim bt1 As Button
    Set bt1 = sht.Buttons.Add(0.6, 15.6, 112.2, 27.6)
    bt1.OnAction = "Pomodoro_Timer.xlsb!PomodoroSession"
    bt1.Text = "Start"
    
    Dim bt2 As Button
    Set bt2 = sht.Buttons.Add(1.8, 56.4, 111, 30)
    bt2.OnAction = "Pomodoro_Timer.xlsb!Clear_all_records"
    bt2.Text = "Clear Records"
    
    Range("A1").Select
    
    '*******************************
    'Summary Sheet
    '*******************************
    Set sht = sht2
    sht.Select
    
    'Column width
    sht.Columns(1).ColumnWidth = 131.4
    sht.Columns(2).ColumnWidth = 87

    'Pivot Table
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "Table24", Version:=6).CreatePivotTable TableDestination:="'" & sht.Name & "'!R1C1", _
    TableName:="PivotTable1", DefaultVersion:=6
    
    With sht.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    sht.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    'Filter
    With sht.PivotTables("PivotTable1").PivotFields("Date")
        .Orientation = xlPageField
        .Position = 1
    End With
    'Row values
    With sht.PivotTables("PivotTable1").PivotFields("Task")
        .Orientation = xlRowField
        .Position = 1
    End With
    'Calculated fields
    sht.PivotTables("PivotTable1").CalculatedFields.Add "Duration", _
        "=End - Start", True
    sht.PivotTables("PivotTable1").AddDataField sht.PivotTables( _
        "PivotTable1").PivotFields("Duration"), "Duration ", xlSum
        
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Duration ")
        .NumberFormat = "hh:mm;@"
    End With
    
    'Buttons
    Dim btx As Button
    Set btx = sht.Buttons.Add(270, 0.6, 191.4, 28.2)
    btx.OnAction = "Pomodoro_Timer.xlsb!Refresh_Summary_PivotTable"
    btx.Text = "Refresh Table"
    
    Range("A1").Select
    
    '*******************************
    'Recent Sheet
    '*******************************
    Set sht = sht3
    sht.Select
    
    'Column width
    sht.Columns(1).ColumnWidth = 140 / 5
    
    'Cell values
    sht.Cells(1, 1).Value2 = "Recent Tasks"
    sht.Cells(2, 1).Value2 = "Check emails"
    sht.Cells(3, 1).Value2 = "Make phone call"
    sht.Cells(4, 1).Value2 = "Reading"

    'Formatting
    With sht.Cells(1, 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12874308
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With sht.Cells(1, 1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    sht.Cells(1, 1).Font.Bold = True

    'Buttons
    Dim bt3 As Button
    Set bt3 = sht.Buttons.Add(270, 0.6, 191.4, 28.2)
    bt3.OnAction = "Pomodoro_Timer.xlsb!Clear_Recent_Tasks"
    bt3.Text = "Clear Recent Task"
    
    Range("A1").Select

    '*******************************
    'Settings Sheet
    '*******************************
    Set sht = sht4
    sht.Select
    
    'Column width
    sht.Columns(1).ColumnWidth = 200 / 5
    sht.Columns(2).ColumnWidth = 55 / 5
    
    'Cell values
    sht.Cells(1, 1).Value2 = "Settings"
    sht.Cells(2, 1).Value2 = "Pomodoro duration (min)"
    sht.Cells(3, 1).Value2 = "Pomodoro duration (sec)"
    sht.Cells(4, 1).Value2 = "Break duration (min)"
    sht.Cells(5, 1).Value2 = "Break duration (sec)"
    sht.Cells(6, 1).Value2 = "Open Timer in a separate Excel instance"
    sht.Cells(7, 1).Value2 = "Reactivate Excel window when timer is closed"
    sht.Cells(8, 1).Value2 = "Record unfinished Pomodoro session"
    sht.Cells(9, 1).Value2 = "Don't record if session was less than (min)"
    sht.Cells(10, 1).Value2 = "Play sound at the end of Pomodoro session"
    sht.Cells(11, 1).Value2 = "Play sound at the end of Break"
    sht.Cells(12, 1).Value2 = "Use custom position"
    sht.Cells(13, 1).Value2 = "Left position"
    sht.Cells(14, 1).Value2 = "Top position"
    sht.Cells(15, 1).Value2 = "Use shortcuts (F10)"
    sht.Cells(16, 1).Value2 = "Flasing color"

    sht.Cells(1, 2).Value2 = "Value"
    sht.Cells(2, 2).Value2 = 25
    sht.Cells(3, 2).Value2 = 0
    sht.Cells(4, 2).Value2 = 5
    sht.Cells(5, 2).Value2 = 0
    sht.Cells(6, 2).Value2 = True
    sht.Cells(7, 2).Value2 = True
    sht.Cells(8, 2).Value2 = True
    sht.Cells(9, 2).Value2 = 1
    sht.Cells(10, 2).Value2 = True
    sht.Cells(11, 2).Value2 = True
    sht.Cells(12, 2).Value2 = False
    sht.Cells(13, 2).Value2 = 0.5
    sht.Cells(14, 2).Value2 = 0.5
    sht.Cells(15, 2).Value2 = True
    sht.Cells(16, 2).Value2 = ""

    'Table
    sht.ListObjects.Add(xlSrcRange, sht.Range("$A$1:$B$16"), , xlYes).Name = "Table2"
    
    'Formatting
    With sht.Range("B16").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16711680
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    sht.Range("B13:B14").Style = "Percent"
    
    'Datavalidation
    sht.Select
    Range("B2").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="=24*60"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="60"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B4").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="=24*60"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B5").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="60"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B6").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B7").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B8").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B9").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="=B2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B10").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B11").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B13").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween, Formula1:="0", Formula2:="100"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B14").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween, Formula1:="0", Formula2:="100"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B15").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    Range("A1").Select
    
    Sheets("Pomodoro").Select
    wb.SaveAs Filename:=ThisWorkbook.Path & "\" & SHORT_NAME & EXT, FileFormat:=xlExcel12
    
End Sub


Function testFileLocation() As Boolean
    
    Dim ErrMsg As String
    
    'Test if this workbook has been saved
    Dim FileEverSaved As Boolean
    If ThisWorkbook.Path = "" Then
        ErrMsg = "Please save the file that contains the Assembler module in the same folder than Installer.bas and try again"
        MsgBox ErrMsg, vbCritical
        testFileLocation = False
        Exit Function
    End If
    
    'Test if the src folder contains a folder with the right name
    Dim SourceFolderExist As Boolean
    If Dir(ThisWorkbook.Path & "\src\" & SHORT_NAME & EXT, vbDirectory) = "" Then
        ErrMsg = "Please save the file that contains the Assembler module in a location where the source folder (src) contains a folder named " & SHORT_NAME & EXT
        MsgBox ErrMsg, vbCritical
        testFileLocation = False
        Exit Function
    End If
    
    testFileLocation = True
    
End Function
