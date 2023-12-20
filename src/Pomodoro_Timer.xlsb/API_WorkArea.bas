Attribute VB_Name = "API_WorkArea"
'PURPOSE: Get screen size in pixels
'REFERENCE: https://www.excelforum.com/excel-programming-vba-macros/565556-why-does-spi_getworkarea-come-in-too-large.html

Option Explicit

Private Const SPI_GETWORKAREA = 48

#If VBA7 Then
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As Any, _
    ByVal fuWinIni As Long) As Long
#Else
    Private Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As Any, _
    ByVal fuWinIni As Long) As Long
#End If

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Function GETWORKAREA_HEIGHT() As Double
'PURPOSE: Get the screen size exluding the taskbar
    Dim nRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, nRect, 0
    GETWORKAREA_HEIGHT = (nRect.Bottom - nRect.Top)
End Function


Function GETWORKAREA_WIDTH() As Double
'PURPOSE: Get the screen size exluding the taskbar
    Dim nRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, nRect, 0
    GETWORKAREA_WIDTH = (nRect.Right - nRect.Left)
End Function


