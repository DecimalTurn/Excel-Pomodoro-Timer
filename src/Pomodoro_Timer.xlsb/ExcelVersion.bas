Attribute VB_Name = "ExcelVersion"
Attribute VB_Description = "Get Excel version informations. REFERENCE: https://www.rondebruin.nl/mac/mac001.htm"
'@Folder("Utilities")
'@ModuleDescription "Get Excel version informations. REFERENCE: https://www.rondebruin.nl/mac/mac001.htm"
Option Explicit


Public Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#Else
    IsMac = False
#End If
End Function

Public Function Is64BitOffice() As Boolean
#If Win64 Then
    Is64BitOffice = True
#End If
End Function

Public Function ExcelVersionNumber() As Double
'Win Excel versions are always a whole number (15)
'Mac Excel versions show also the number of the update (15.29)
    ExcelVersionNumber = Val(Application.Version)
End Function
