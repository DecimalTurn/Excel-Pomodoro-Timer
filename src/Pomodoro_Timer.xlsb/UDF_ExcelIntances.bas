Attribute VB_Name = "UDF_ExcelIntances"
'PURPOSE: The functions in this module are used to calculate the number of Excel instances currently open

Option Explicit



Function ExcelInstances()
    ExcelInstances = Arr_Countif(AllRunningApps, "EXCEL.EXE")
End Function

Public Function AllRunningApps() As Variant
    'Reference: http://www.vbaexpress.com/forum/archive/index.php/t-36677.html
    Dim strComputer As String
    Dim objServices As Object, objProcessSet As Object, Process As Object
    Dim oDic As Object, a() As Variant
    Dim i As Integer
    Set oDic = CreateObject("Scripting.Dictionary")
    strComputer = "."
    Set objServices = GetObject("winmgmts:\\" _
    & strComputer & "\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery _
    ("SELECT Name FROM Win32_Process", , 48)
    For Each Process In objProcessSet
    i = i + 1
    ReDim Preserve a(1 To i)
    a(i) = Process.Name
    Next
    Set objProcessSet = Nothing
    Set oDic = Nothing
    AllRunningApps = a()
End Function

Private Function Arr_Countif(arr, criteria) As Long
    Dim i As Long, el As Variant
    If TypeName(el) = "string" Then
        For Each el In arr
            If UCase(el) = UCase(criteria) Then
                i = i + 1
            End If
        Next el
    Else
        For Each el In arr
            If el = criteria Then
                i = i + 1
            End If
        Next el
    End If
    Arr_Countif = i
End Function


