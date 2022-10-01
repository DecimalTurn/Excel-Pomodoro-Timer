VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PomodoroTimer 
   Caption         =   "Timer"
   ClientHeight    =   924
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2190
   OleObjectBlob   =   "PomodoroTimer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PomodoroTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Countdown timer
'REFERENCE: https://www.mrexcel.com/forum/excel-questions/594922-countdown-timer-userform.html

Option Explicit

Const sleeptime = 10 'Miliseconds

Private Sub UserForm_Initialize()
    UFIsVisible = True
    'Position of the Userform
    If ThisWorkbook.Sheets("Settings").Range("Custom_position") = True And Not IsMac Then
        Me.StartUpPosition = 0
        Me.Top = ThisWorkbook.Sheets("Settings").Range("Top_pos").Value2 * (PointPerPixelY() * GETWORKAREA_HEIGHT - Me.Height)
        Me.Left = ThisWorkbook.Sheets("Settings").Range("Left_pos").Value2 * (PointPerPixelX() * GETWORKAREA_WIDTH - Me.Width)
    ElseIf Not IsMac Then
        'Reposition the window
        Me.StartUpPosition = 0
        Me.Top = PointPerPixelY() * GETWORKAREA_HEIGHT - Me.Height
        Me.Left = PointPerPixelX() * GETWORKAREA_WIDTH - Me.Width
    End If
    OngoingTimer = False

    Dim M As Double, S As Double
    M = Int(AllowedTime)
    S = AllowedTimeSec
    
    With tBx1
        .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
    End With

    
    'The code below makes sure that the userform stays on top of all windows.
    'Source: https://www.mrexcel.com/forum/excel-questions/386643-userform-always-top.html
    If Not IsMac Then
        AlwaysOnTop Me.caption
    End If
    
    If AutoLaunch Then
        If IsMac Then
            Call Launch_timer_mac
        End If
    End If

End Sub
Private Sub UserForm_Activate()
        If AutoLaunch Then
            If Not IsMac Then
                Call Launch_timer
            End If
        End If
End Sub

Private Sub Launch_timer()
        
    Dim calc_iniset As Variant: calc_iniset = Application.Calculation
    Call Optimize_VBA_Performance(True)
    
    OngoingTimer = True
    StopTimer = False
    CloseTimer = False
    CommandButton2.caption = "Cancel"
    
    'Reset the colors
    PomodoroTimer.BackColor = -2147483633
    TextBox2.BackColor = -2147483633
    tBx1.BackColor = -2147483633
    
    StartTime = Now()
    TodaysDate = Date
    
    
    Dim M As Double, S As Double
    Dim TotalTime
    Dim EndTime As Double
    Dim RemaingTime As Double
    
    TotalTime = 60 * AllowedTime + AllowedTimeSec
    EndTime = DateAdd("s", TotalTime, Now())
    RemaingTime = DateDiff("s", Now(), EndTime)
    
    
    Do
        
        RemaingTime = DateDiff("s", Now(), EndTime)
        M = Int(RemaingTime / 60)
        S = RemaingTime - 60 * M
        
        With tBx1
            .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
        End With
        
        'Released the control to the OS
        DoEvents
        
        'Now sleep for 0.1 sec
        Call Sleep(sleeptime)
        
        'Stop the code if the form is not visible
        If UFIsVisible = False Then: Debug.Print "Form is not visible. The code will now stop.": End
        
    Loop Until RemaingTime <= 0 Or StopTimer
    
    'Recording session
    If StopTimer = False Or ThisWorkbook.Sheets("Settings").Range("Record_unfinished").Value2 = True Then
        If (TotalTime - RemaingTime) / 60 > ThisWorkbook.Sheets("Settings").Range("No_Recording_limit") Then
            Call Add_new_record(TodaysDate, StartTime, Now, Not (StopTimer), ThisWorkbook.Sheets("Pomodoro").Range("TaskNameRng"))
        End If
    End If
    
    Call Optimize_VBA_Performance(False, calc_iniset)
    
    If StopTimer = False Then 'If the timer was stopped by the user
        'Proceed with the Break
        If ThisWorkbook.Sheets("Settings").Range("Sound_end_Pomodoro") = True Then Beep
        TextBox2.Value = "Break"
        Call TakeBreak
    Else
        'Do nothing
        CommandButton2.caption = "Start"
        OngoingTimer = False
    End If
    
    If CloseTimer Then Unload Me

End Sub

Private Sub TakeBreak()
    'Reset StopTimer:
    StopTimer = False
    
    Dim calc_iniset As Variant: calc_iniset = Application.Calculation
    Call Optimize_VBA_Performance(True)
    
    Dim M As Double, S As Double
    M = BreakTime
    S = BreakTimeSec
     
    With tBx1
        .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
    End With
    
    Call Optimize_VBA_Performance(False, calc_iniset)
    
    Call TakeBreak2

End Sub

Private Sub TakeBreak2()
    Dim M As Long, S As Long
    Dim EndTime As Double
    Dim RemaingTime As Double
    Dim TotalTime As Long
    
    TotalTime = 60 * BreakTime + BreakTimeSec
    EndTime = DateAdd("s", TotalTime, Now())
    RemaingTime = DateDiff("s", Now(), EndTime)
    
    
    Do
        RemaingTime = DateDiff("s", Now(), EndTime)
        M = Int(RemaingTime / 60)
        S = RemaingTime - 60 * M
        
        'Flashing
        If TotalTime - RemaingTime < 9 Then
            If S Mod 2 = 1 Then
                PomodoroTimer.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
                TextBox2.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
                tBx1.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
            Else
                PomodoroTimer.BackColor = -2147483633 'Normal color
                TextBox2.BackColor = -2147483633 'Normal color
                tBx1.BackColor = -2147483633 'Normal color
            End If
        End If
        
        With tBx1
            .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
        End With
        'Released the control to the OS
        DoEvents
        'Now sleep for 0.1 sec
        Call Sleep(sleeptime)
    Loop Until RemaingTime <= 0 Or StopTimer

    If StopTimer = False Then
        If ThisWorkbook.Sheets("Settings").Range("Sound_end_Break") = True Then Beep
        'Remain in color to get the user's attention
        PomodoroTimer.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
        TextBox2.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
        tBx1.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
    Else
        PomodoroTimer.BackColor = -2147483633 'Normal color
        TextBox2.BackColor = -2147483633 'Normal color
        tBx1.BackColor = -2147483633 'Normal color
    End If
        TextBox2.Value = ""
        CommandButton2.caption = "Start"
        OngoingTimer = False
            
        'Redo basic calculations form the initialize macro
        M = Int(AllowedTime)
        S = (AllowedTime - Int(AllowedTime)) * 60
         With tBx1
            .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
        End With
    
        
    
End Sub

Private Sub CommandButton2_Click()
If OngoingTimer = False Then 'Start the timer
    UFIsVisible = True 'The form must be visible
    ThisWorkbook.Application.WindowState = xlMinimized
    CommandButton2.caption = "Cancel"
    If Not IsMac Then
        Call Launch_timer
    Else
        Call Launch_timer_mac
    End If
Else 'Stop the timer
    StopTimer = True
    OngoingTimer = False
    'No need to unload the userform here since the main procedure (Launch_timer) will take care of that as long as we are still in the loop
    'Unload Me
End If
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'PURPOSE: This procedure will run if the user click on the "X" to close the userform.
    Dim Wkb As Workbook
    
    Set Wkb = ThisWorkbook
    StopTimer = True
    CloseTimer = True
    
    'At this point, since the user clicked on the userform to close it. Excel is the active window, but it might not be on top.
    'Make Excel the active window (optional)
    On Error Resume Next
    If ThisWorkbook.Sheets("Settings").Range("Reopen_Excel_after_x").Value2 = True And Not IsMac Then
        Call AppActivate(Wkb.Application.caption, True)
        ShowWindow GetForegroundWindow, SW_SHOWMAXIMIZED
    End If
    On Error GoTo 0
    
    'Make sure that the form isn't considered visible anymore
    UFIsVisible = False
    
End Sub


Private Sub AlwaysOnTop(caption As String)
'PURPOSE: This function allows the userform to remain on top of all windows - Adjusted
'REFERENCE: https://www.mrexcel.com/forum/excel-questions/386643-userform-always-top-2.html

    #If VBA7 Then
        Dim hWnd As LongPtr
    #Else
        Dim hWnd As Long
    #End If
    Dim lResult As Boolean
    
    If Val(Application.Version) >= 9 Then
        hWnd = FindWindow("ThunderDFrame", caption)
    Else
        hWnd = FindWindow("ThunderXFrame", caption)
    End If
    
    If hWnd <> 0 Then
    
        lResult = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        
    Else
    
        MsgBox "AlwaysOnTop: userform with caption '" & caption & "' not found"
        
    End If
    
End Sub

Private Sub UserForm_Terminate()
    UFIsVisible = False
End Sub
