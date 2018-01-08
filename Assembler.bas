Attribute VB_Name = "Assembler"
Option Explicit

Sub install_vbaDevelopper()

    Dim aWb As Workbook
    On Error Resume Next
    Set aWb = Workbooks("vbadeveloper.xlam")
    On Error GoTo 0
    If aWb Is Nothing Then
        Dim instWb As Workbook
        Set instWb = Workbooks.Open(ThisWorkbook.Path & "/vbaDeveloper/Installer.xls")
        Application.Run "installer.xls!AutoAssembler"
    Else
        Dim objExcel As Excel.Application
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = True
        'objExcel.Workbooks.Add
        Set aWb = objExcel.Workbooks.Open(ThisWorkbook.Path & "/vbaDeveloper/Installer.xls")
        objExcel.Application.Run aWb.Name & "!AutoAssembler"
        'aWb.Close (False)
        'objExcel.Quit
    End If

End Sub

Sub Assemble()
    
    Dim aWb As Workbook, twb As Workbook
    Set twb = ThisWorkbook
    On Error Resume Next
    Set aWb = Workbooks("vbadeveloper.xlam")
    On Error GoTo 0
    If aWb Is Nothing Then
        Set aWb = Workbooks.Open(ThisWorkbook.Path & "/vbaDeveloper/vbadeveloper.xlam")
        Application.Run "'" & aWb.Name & "'!ImportVbaAndXMLCode", aWb.Path & "\src\Pomodoro_Timer.xlsb"
    Else
        Dim objExcel As Excel.Application
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = True
        objExcel.Workbooks.Add
        Set aWb = objExcel.Workbooks.Open(ThisWorkbook.Path & "\vbaDeveloper\vbaDeveloper.xlam")
        objExcel.Application.Run aWb.Name & "!ImportVbaAndXMLCode", twb.Path & "\src\Pomodoro_Timer.xlsb"
        'Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 8)
        'aWb.Close (False)
        'objExcel.Quit
    End If

End Sub
