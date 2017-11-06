Attribute VB_Name = "OnedriveExport"
Option Explicit

Private Sub FSO_Copy_paste()
'PURPOSE: Copy-paste this file to Server

Dim objFSO As Object 'Late binding
Set objFSO = CreateObject("Scripting.FileSystemObject") 'Create an instance of the FileSystemObject
Dim Onedrivepath As String
Onedrivepath = "C:\Users\Martin\OneDrive\Programs\VBA\PERSONAL\Projects"

objFSO.CopyFile ThisWorkbook.FullName, Onedrivepath & "\" & ThisWorkbook.Name

MsgBox "The upload was successful."

End Sub
