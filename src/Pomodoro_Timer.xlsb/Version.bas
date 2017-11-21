Attribute VB_Name = "Version"
Option Explicit
'REFERENCE: https://www.rondebruin.nl/mac/mac001.htm

Public Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#End If
End Function

Public Function Is64BitOffice() As Boolean
#If Win64 Then
    Is64BitOffice = True
#End If
End Function

Public Function Excelversion() As Double
'Win Excel versions are always a whole number (15)
'Mac Excel versions show also the number of the update (15.29)
    Excelversion = Val(Application.Version)
End Function
