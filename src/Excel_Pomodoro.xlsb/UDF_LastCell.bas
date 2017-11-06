Attribute VB_Name = "UDF_LastCell"
Option Explicit

Function LastCell_row(sht As Worksheet) As Long
'PURPOSE: 'Calculate the last row of data based on column A.

    Dim tst As Range, rownb As Long, oldnb As Long
    Set tst = sht.Range("A1")
    rownb = tst.Row
    Do
        oldnb = rownb
        Set tst = tst.End(xlDown)
        rownb = tst.Row
    Loop While rownb > oldnb
    
If Not IsEmpty(sht.Cells(rownb, 1)) Then
    LastCell_row = rownb
Else
    LastCell_row = sht.Cells(rownb, 1).End(xlUp).Row
End If

End Function
