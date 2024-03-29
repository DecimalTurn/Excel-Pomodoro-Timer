Attribute VB_Name = "API_Sleep"
'PURPOSE: Define the sleep function to stop the code from running and releasing CPU usage.

Option Explicit

#If VBA7 Then ' Excel 2010 or later for Windows

    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 64 Bit Systems

#Else ' pre Excel 2010 for Windows
    
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems

#End If


Sub SleepTest()
'MsgBox "Execution is started"
Sleep 10000 'delay in milliseconds
MsgBox "Waiting completed"
End Sub
