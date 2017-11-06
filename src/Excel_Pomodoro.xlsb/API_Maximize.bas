Attribute VB_Name = "API_Maximize"
'PURPOSE: Contain function that allows to maximize or minimize a window.
'REFERENCE: http://www.vbaexpress.com/forum/archive/index.php/t-36677.html

Option Explicit

#If VBA7 Then
   
    Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
#Else
    Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Declare Function GetForegroundWindow Lib "user32" () As Long
#End If

' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10


