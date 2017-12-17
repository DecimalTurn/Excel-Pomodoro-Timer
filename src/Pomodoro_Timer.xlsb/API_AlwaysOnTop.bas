Attribute VB_Name = "API_AlwaysOnTop"
'REFERENCE: https://www.mrexcel.com/forum/excel-questions/386643-userform-always-top-2.html
'PURPOSE: This module includes the functions used to make sure that the Timer stays on top of all windows.

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

'https://msdn.microsoft.com/en-us/library/office/gg264421.aspx
'64-Bit Visual Basic for Applications Overview
'See also: https://sysmod.wordpress.com/2016/09/03/conditional-compilation-vba-excel-macwin3264/
'For Mac declarations


#If VBA7 Then ' Excel 2010 or later for Windows

    'VBA version 7 compiler, therefore >= Office 2010
    'PtrSafe means function works in 32-bit and 64-bit Office
    'LongPtr type alias resolves to Long (32 bits) in 32-bit Office, or LongLong (64 bits) in 64-bit Office

    Public Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hWnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr

#Else ' pre Excel 2010 for Windows
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

#End If



