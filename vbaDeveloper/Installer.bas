Attribute VB_Name = "Installer"
Option Explicit

'1) Create an Excel file called Installer.xlsm in same folder than Installer.bas:
'   *\GIT\vbaDeveloper-master\

'2) Open the VB Editor (Alt+F11) right click on the active project and choose Import a file and chose:
'    *\GIT\vbaDeveloper-master\Installer.bas


'3a) Go in Tools--> References and activate:
'   - Microsoft Scripting Runtime
'   - Microsoft Visual Basic for Application Extensibility X.X

'3b) Enable programatic access to VBA:
'       File -> Options -> Trust Center, Trust Center Settings, -> Macros,
'       tick the box: 'Enable programatic access to VBA'  (In excel 2010: 'Trust access to the vba project object model')

'4) Run the Sub AutoInstaller in the module Installer


'5) Create a new excel file and also open the file vbaDeveloper.xlam located in the folder: *\GIT\vbaDeveloper-master\

'6) Make step 3a and 3b again for this file and run the sub testImport located in the module "Build".

Public Const TOOL_NAME = "vbaDeveloper"

Sub AutoInstaller()

    AutoInstaller_step0

End Sub
Sub AutoInstaller_step0()

    'Close the vbaDevelopper Workbook if already open and uninstall from addins
    On Error Resume Next
    Workbooks(TOOL_NAME & ".xlam").Close
    Application.AddIns2(AddinName2index(TOOL_NAME & ".xlam")).Installed = False
    On Error GoTo 0

    Application.OnTime Now + TimeValue("00:00:06"), "AutoInstaller_step1"
    
End Sub

Sub AutoInstaller_step1()

'Prepare variable
Dim CurrentWB As Workbook, NewWB As Workbook

Dim textline As String, strPathOfBuild As String, strLocationXLAM As String

'Set the variables
Set CurrentWB = ThisWorkbook
Set NewWB = Workbooks.Add

'Import code form Build.bas  to the new workbook
strPathOfBuild = CurrentWB.Path & "\src\vbaDeveloper.xlam\Build.bas"
NewWB.VBProject.VBComponents.Import strPathOfBuild

    'Rename the project (in the VBA) to vbaDeveloper
    NewWB.VBProject.Name = TOOL_NAME

    'Add references to the library
        'Microsoft Scripting Runtime
        NewWB.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
        
        'Microsoft Visual Basic for Applications Extensibility 5.3
        NewWB.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    
    'In VB Editor, press F4, then under Microsoft Excel Objects, select ThisWorkbook.Set the property 'IsAddin' to TRUE
    NewWB.IsAddin = True
    'In VB Editor, menu File-->Save Book1; Save as vbaDeveloper.xlam in the same directory as 'src'

    strLocationXLAM = CurrentWB.Path
    NewWB.SaveAs strLocationXLAM & "\" & TOOL_NAME & ".xlam", xlOpenXMLAddIn
        
    'Close excel. Open excel with a new workbook, then open the just saved vbaDeveloper.xlam
    NewWB.Close savechanges:=False
    
    'Add the Add-in (if not already present)
    If IsAddinInstalled(TOOL_NAME & ".xlam") = False Then
        Call Application.AddIns2.Add(strLocationXLAM & "\" & TOOL_NAME & ".xlam", CopyFile:=False)
    End If
    
    'Continue to step 2
    Application.OnTime Now + TimeValue("00:00:02"), "AutoInstaller_step2"
    
End Sub

Sub AutoInstaller_step2()

    'Install the Addin (This should open the file)
    Application.AddIns2(AddinName2index(TOOL_NAME & ".xlam")).Installed = True
    
    Application.OnTime Now + TimeValue("00:00:02"), "AutoInstaller_step3"
    
End Sub


Sub AutoInstaller_step3()

    'Run the Build macro in vbaDeveloper
    Application.Run "vbaDeveloper.xlam!Build.testImport"

    'Continue to step 4
    Application.OnTime Now + TimeValue("00:00:06"), "AutoInstaller_step4"
    
End Sub

Sub AutoInstaller_step4()

    'Run the Workbook_Open macro from vbaDeveloper
    Application.Run "vbaDeveloper.xlam!Menu.createMenu"
    
    Workbooks(TOOL_NAME & ".xlam").Save
    
    MsgBox TOOL_NAME & " was successfully installed."
    
End Sub

Function IsAddinInstalled(ByVal addin_name As String) As Boolean
'PURPOSE: Return true if the Add-in is installed
    If AddinName2index(addin_name) > 0 Then
        IsAddinInstalled = True
    ElseIf AddinName2index(addin_name) = 0 Then
        IsAddinInstalled = False
    End If
End Function

Function AddinName2index(ByVal addin_name As String) As Integer
'PURPOSE: Convert the name of an installed addin to its index
    Dim i As Variant
    For i = 1 To Excel.Application.AddIns2.Count
        If Excel.Application.AddIns2(i).Name = addin_name Then
            AddinName2index = i
            Exit Function
        End If
    Next
    'If we get to this line, it means no match was found
    AddinName2index = 0
End Function

Sub AutoAssembler()

'Prepare variable
Dim CurrentWB As Workbook, NewWB As Workbook

Dim textline As String, strPathOfBuild As String, strLocationXLAM As String

'Set the variables
Set CurrentWB = ThisWorkbook
Set NewWB = Workbooks.Add

'Import code form Build.bas  to the new workbook
strPathOfBuild = CurrentWB.Path & "\src\vbaDeveloper.xlam\Build.bas"
NewWB.VBProject.VBComponents.Import strPathOfBuild

    'Rename the project (in the VBA) to vbaDeveloper
    NewWB.VBProject.Name = TOOL_NAME

    'Add references to the library
        'Microsoft Scripting Runtime
        NewWB.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
        
        'Microsoft Visual Basic for Applications Extensibility 5.3
        NewWB.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    
    'In VB Editor, press F4, then under Microsoft Excel Objects, select ThisWorkbook.Set the property 'IsAddin' to TRUE
    NewWB.IsAddin = True
    'In VB Editor, menu File-->Save Book1; Save as vbaDeveloper.xlam in the same directory as 'src'
    
    'Save file as .xlam
    strLocationXLAM = CurrentWB.Path
    Application.DisplayAlerts = False
    NewWB.SaveAs strLocationXLAM & "\" & TOOL_NAME & ".xlam", xlOpenXMLAddIn
    Application.DisplayAlerts = True
    
    'Close and reopen the file
    NewWB.Close savechanges:=False
    Set NewWB = Workbooks.Open(strLocationXLAM & "\" & TOOL_NAME & ".xlam")
    
    Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 5)
    
    'Run the Build macro in vbaDeveloper
    Application.OnTime Now + TimeValue("00:00:05"), "vbaDeveloper.xlam!Build.testImport"
    Application.OnTime Now + TimeValue("00:00:12"), "installer.xls!SaveFile"
    
End Sub

Sub SaveFile()
    Workbooks("vbaDeveloper.xlam").Save
    ThisWorkbook.Saved = True
    Application.Quit
End Sub
