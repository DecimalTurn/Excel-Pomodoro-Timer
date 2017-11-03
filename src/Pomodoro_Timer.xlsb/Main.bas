Attribute VB_Name = "Main"
'Infos: different modes for userform:
'https://www.mrexcel.com/forum/excel-questions/465425-minimize-excel-leave-userform-showing.html


Public AllowedTime As Integer 'Number of minutes to count down
Public AllowedTimeSec  As Integer 'Number of seconds to count down
Public BreakTime As Double
Public BreakTimeSec As Integer
Public AutoLaunch As Boolean
Public TaskName As String
Public StopTimer As Boolean 'User stopped timer
Public CloseTimer As Boolean 'User clicked the X
Public OngoingTimer As Boolean  'Take the value true after the timer has started (was initialized)
Public StartTime As Variant
Public TodaysDate As Variant

Sub PomodoroSession()
    AllowedTime = Range("Pomodoro")
    AllowedTimeSec = Range("Pomodoro_sec")
    BreakTime = Range("Break")
    BreakTimeSec = Range("Break_sec")
    AutoLaunch = True
    ThisWorkbook.Application.WindowState = xlMinimized
    If Reopen_decision = True Then
        MsgBox "To let you work with Excel while the timer is running, this file will now be reopen in a second instance of Excel." & vbNewLine & _
        "Once, the was has been reopen, you will need to relaunch the timer."
        Call OpenItSelfInAnotherInstance
    End If
    PomodoroTimer.Show vbModeless
    'Note:vbModeless as opposed to vbModal will allow the Excel application to be unlocked while the timer is running
End Sub


