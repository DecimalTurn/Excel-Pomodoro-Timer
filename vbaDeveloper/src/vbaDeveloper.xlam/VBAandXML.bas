Attribute VB_Name = "VBAandXML"
Option Explicit

Sub exportVbaAndXMLCode()

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    wb.Save

    Call Build.exportVbaCode(wb.VBProject)
    Call XMLexporter.unpackXML(wb.name)
    
    'Delete the VBProject bin file
    'On Error Resume Next
    'Delete files
    Dim FSO As New Scripting.FileSystemObject
    FSO.DeleteFile wb.Path & "\src\" & wb.name & "\XMLsource\xl\vbaProject.bin", True
    'On Error GoTo 0
    
    MsgBox "Successfully exported VB code and XML content."
End Sub

Sub ImportVbaAndXMLCode_ActiveWorkbook()

    Call ImportVbaAndXMLCode

End Sub

Sub ImportVbaAndXMLCode(Optional ByVal FileFolderPath As String)
    
    Dim wb As Workbook
    Dim oFileName As String, nFileName As String, nShortFileName As String, oFileFolderPath As String
    
    If FileFolderPath = vbNullString Then
        Set wb = ActiveWorkbook
        oFileFolderPath = wb.Path
        oFileName = wb.FullName
        nShortFileName = wb.name
        wb.Close
    Else
        oFileFolderPath = Left(FileFolderPath, InStr(FileFolderPath, "\src") - 1)
        oFileName = Replace(FileFolderPath, "\src", "")
        nShortFileName = Split(FileFolderPath, "\")(UBound(Split(FileFolderPath, "\")))
    End If
    
    Dim nwb As Workbook
    Dim ErrFlag  As Boolean, ErrMsg As String
    ErrFlag = False
    ErrMsg = ""
    
    'Ask the user to confirm
    Dim ireply As Variant
    ireply = MsgBox(prompt:="Are you sure that you want to overwrite " & nShortFileName, Buttons:=vbYesNo, title:="Decision")
    
    If ireply = vbYes Then
        'Do nothing (Continue)
    ElseIf ireply = vbNo Then
        Exit Sub
    Else 'They cancelled (VbCancel)
        Exit Sub
    End If
    
    Call XMLexporter.rebuildXML(oFileFolderPath, oFileFolderPath & "\src\" & nShortFileName, ErrFlag, ErrMsg, nwb)
    
    If ErrFlag = True Then
        MsgBox (ErrMsg)
        Exit Sub
    End If
    
    nFileName = nwb.FullName
    nwb.Close
    
    Dim FSO As New Scripting.FileSystemObject
    Dim nFile As file, oFile As file
    Set oFile = FSO.GetFile(oFileName)
    Set nFile = FSO.GetFile(nFileName)
    
    oFile.Delete
    nFile.name = nShortFileName
    
    Set nwb = Workbooks.Open(oFileName)
    
    Call Build.importVbaCode(nwb.VBProject)
    MsgBox "Successfully imported VB code and XML content."
    
End Sub
