Attribute VB_Name = "API_AlwaysOnTop"
'PURPOSE: This module includes the functions used to make sure that the Timer stays on top of all windows.
'REFERENCE: https://www.mrexcel.com/forum/excel-questions/386643-userform-always-top-2.html

Option Explicit

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

' For hWndInsertAfter in SetWindowPos
Public Enum HWND_TYPE
    HWND_TOP = 0
    HWND_NOTOPMOST = -2
    HWND_TOPMOST = -1
    HWND_BOTTOM = 1
End Enum

' For nIndex in SetWindowLongPtr
Public Enum GWL_TYPE
    GWL_EXSTYLE = -20
    GWL_STYLE = -16
    GWLP_HINSTANCE = -6
    GWLP_ID = -12
    GWLP_USERDATA = -21
    GWLP_WNDPROC = -4
End Enum

' For dwNewLong in SetWindowLongPtr
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000

'https://msdn.microsoft.com/en-us/library/office/gg264421.aspx
'64-Bit Visual Basic for Applications Overview

#If VBA7 Then

    'VBA version 7 compiler, therefore >= Office 2010
    'PtrSafe means function works in 32-bit and 64-bit Office
    'LongPtr type alias resolves to Long (32 bits) in 32-bit Office, or LongLong (64 bits) in 64-bit Office

    Public Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hWnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal x As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
        
    'GetWindowLongPtr (Uses different alias (true name) between 32-bit and 64-bit)
    #If Win64 Then
        '64-bit Office
        Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
            (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) As LongPtr
    #Else
        '32-bit Office
        Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" _
            (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) As LongPtr
    #End If
    
    'Set WindowsLongPtr (Uses different alias (true name) between 32-bit and 64-bit)
    #If Win64 Then
        Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" _
            (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long, _
            ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" _
            (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long, _
            ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
        (ByVal hWnd As LongPtr) As Long

#Else
    
    'VBA version 6 or earlier compiler, therefore <= Office 2007
    
    Public Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long

    Public Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, _
        ByVal nIndex As Long) As Long
    
    Public Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

    Public Declare Function DrawMenuBar Lib "user32" _
        (ByVal hWnd As Long) As Long

#End If
