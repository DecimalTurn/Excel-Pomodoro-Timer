Attribute VB_Name = "winPixel"
'PURPOSE: This module have functions to help convert pixels to points in Excel, allowing to scale things.
'REFERENCE: http://www.vbaexpress.com/forum/showthread.php?21896-Pixel-to-Point-ratio

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
 
Function PointPerPixelX() As Double
    #If VBA7 Then
        Dim hdc As LongPtr
    #Else
        Dim hdc As Long
    #End If
    hdc = GetDC(0)
    PointPerPixelX = 1 / (GetDeviceCaps(hdc, LOGPIXELSX) / 72)
End Function

Function PointPerPixelY() As Double
    #If VBA7 Then
        Dim hdc As LongPtr
    #Else
        Dim hdc As Long
    #End If
    hdc = GetDC(0)
    PointPerPixelY = 1 / (GetDeviceCaps(hdc, LOGPIXELSY) / 72)
End Function

Sub Example()
    #If VBA7 Then
        Dim hdc As LongPtr
    #Else
        Dim hdc As Long
    #End If
    Dim PixPerInchX As Long
    Dim PixPerInchY As Long
    Dim PixPerPtX As Double
    Dim PixPerPtY As Double
    Dim PtPerPixX As Double
    Dim PtPerPixY As Double
     
    hdc = GetDC(0)
     
    PixPerInchX = GetDeviceCaps(hdc, LOGPIXELSX)
    PixPerInchY = GetDeviceCaps(hdc, LOGPIXELSY)
     
    'there are 72 points per inch
    PixPerPtX = PixPerInchX / 72
    PixPerPtY = PixPerInchY / 72
          
    Debug.Print "PixPerPtX:  " & PixPerPtX, "PixPerPtY:  " & PixPerPtY
     
    PtPerPixX = 1 / PixPerPtX
    PtPerPixY = 1 / PixPerPtY
     
    Debug.Print "PtPerPixX:  " & PtPerPixX, "PtPerPixY:  " & PtPerPixX
    ReleaseDC 0, hdc
End Sub

