Attribute VB_Name = "CountDown_mac"
Option Explicit

Const FREQ = 1

Sub Launch_timer_mac()
    'Stop the code if the form is not visible
    If UFIsVisible = False Then: Debug.Print "Form is not visible. The code will now stop.": End
    
    Dim frm As UserForm
    Set frm = PomodoroTimer
    Call Optimize_VBA_Performance(True)
    
    OngoingTimer = True
    StopTimer = False
    CloseTimer = False
    frm.CommandButton2.caption = "Cancel"
    
    'Reset the colors
    PomodoroTimer.BackColor = -2147483633
    frm.TextBox2.BackColor = -2147483633
    frm.tBx1.BackColor = -2147483633
    
    StartTime = Now()
    TodaysDate = Date
    
    
    Dim M As Double, S As Double
    Dim TotalTime
    Dim EndTime As Double
    Dim RemaingTime As Double
    
    TotalTime = 60 * AllowedTime + AllowedTimeSec
    EndTime = DateAdd("s", TotalTime, Now())
    RemaingTime = DateDiff("s", Now(), EndTime)
    
        RemaingTime = DateDiff("s", Now(), EndTime)
        M = Int(RemaingTime / 60)
        S = RemaingTime - 60 * M
        
        With frm.tBx1
            .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
        End With
        
        'Released the control to the OS
        'DoEvents
        
        'Now "sleep"
        Application.OnTime Now + TimeValue("00:00:01") * FREQ, "Launch_timer_mac2"

End Sub

Sub Launch_timer_mac2()
    Dim frm As UserForm
    Set frm = PomodoroTimer
    
    Dim M As Double, S As Double
    Dim TotalTime
    Dim EllapsedtTime
    Dim StartTime As Double
    Dim EndTime As Double
    Dim RemaingTime As Double
    
    TotalTime = 60 * AllowedTime + AllowedTimeSec
'    EllapsedtTime = TotalTime - (60 * Split(frm.tBx1.Value, ":")(0) + 1 * Split(frm.tBx1.Value, ":")(1))
'        M = Int(EllapsedtTime / 60)
'        S = EllapsedtTime - 60 * M
'    EndTime = DateAdd("s", TotalTime, Now())
'    StartTime = Now() - TimeValue("00:" & Format(CStr(M), "00") & ":" & Format(CStr(S), "00"))
    RemaingTime = 60 * Split(frm.tBx1.Value, ":")(0) + 1 * Split(frm.tBx1.Value, ":")(1)
    
    If RemaingTime > 0 And Not StopTimer Then
        RemaingTime = RemaingTime - FREQ
        M = Int(RemaingTime / 60)
        S = RemaingTime - 60 * M
        
        With frm.tBx1
            .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
        End With
        
        'Released the control to the OS
        'DoEvents
        
        'Now "sleep"
        Application.OnTime Now + TimeValue("00:00:01") * FREQ, "Launch_timer_mac2"
        
    Else
    
        'Since we are using the "Application.ontime" technique, it is possible that some public variables will have lost their values
        If TodaysDate = 0 Then TodaysDate = Date
        If StartTime = 0 Then
            EllapsedtTime = TotalTime - RemaingTime
            M = Int(EllapsedtTime / 60)
            S = EllapsedtTime - 60 * M
            StartTime = Now() - TimeValue("00:" & Format(CStr(M), "00") & ":" & Format(CStr(S), "00"))
        End If
        
        'Recording session
        If StopTimer = False Or ThisWorkbook.Sheets("Settings").Range("Record_unfinished").Value2 = True Then
            If (TotalTime - RemaingTime) / 60 > ThisWorkbook.Sheets("Settings").Range("No_Recording_limit") Then
                Call Add_new_record(TodaysDate, StartTime, Now, Not (StopTimer), Range("TaskNameRng"))
            End If
        End If
        
        Call Optimize_VBA_Performance(False, xlAutomatic)
        
        If StopTimer = False Then 'If the timer was stopped by the user
            'Proceed with the Break
            If ThisWorkbook.Sheets("Settings").Range("Sound_end_Pomodoro") = True Then Beep
            frm.TextBox2.Value = "Break"
            Call TakeBreak_mac
        Else
            'Do nothing
            frm.CommandButton2.caption = "Start"
            OngoingTimer = False
        End If
        
        If CloseTimer Then Unload frm
    End If
End Sub

Private Sub TakeBreak_mac()
    Dim frm As UserForm
    Set frm = PomodoroTimer
    'Reset StopTimer:
    StopTimer = False
    
    Call Optimize_VBA_Performance(True)
    
    Dim M As Double, S As Double
    M = BreakTime
    S = BreakTimeSec
     
    With frm.tBx1
        .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
    End With

    Call TakeBreak_mac2

End Sub

Private Sub TakeBreak_mac2()
    Dim frm As UserForm
    Set frm = PomodoroTimer
    Dim M As Long, S As Long
    Dim EndTime As Double
    Dim RemaingTime As Double
    Dim TotalTime As Long
    
    TotalTime = 60 * BreakTime + BreakTimeSec
    RemaingTime = 60 * Split(frm.tBx1.Value, ":")(0) + 1 * Split(frm.tBx1.Value, ":")(1)
    
   If RemaingTime > 0 And Not StopTimer Then
        RemaingTime = RemaingTime - FREQ
        M = Int(RemaingTime / 60)
        S = RemaingTime - 60 * M
        
        'Flashing
        If TotalTime - RemaingTime < 9 Then
            If S Mod 2 = 1 Then
                frm.BackColor = GetRGBColor_Fill(Range("Flashing_color")) 'Flashing color
                frm.TextBox2.BackColor = GetRGBColor_Fill(Range("Flashing_color")) 'Flashing color
                frm.tBx1.BackColor = GetRGBColor_Fill(Range("Flashing_color")) 'Flashing color
            Else
                frm.BackColor = -2147483633 'Normal color
                frm.TextBox2.BackColor = -2147483633 'Normal color
                frm.tBx1.BackColor = -2147483633 'Normal color
            End If
        End If
        
        With frm.tBx1
            .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
        End With
        'Released the control to the OS
        'DoEvents
        'Now "sleep"
        Application.OnTime Now + TimeValue("00:00:01") * FREQ, "TakeBreak_mac2"
        
    Else

        If StopTimer = False Then
            If ThisWorkbook.Sheets("Settings").Range("Sound_end_Break") = True Then Beep
            'Remain in color to get the user's attention
            frm.BackColor = GetRGBColor_Fill(Range("Flashing_color")) 'Flashing color
            frm.TextBox2.BackColor = GetRGBColor_Fill(Range("Flashing_color")) 'Flashing color
            frm.tBx1.BackColor = GetRGBColor_Fill(Range("Flashing_color")) 'Flashing color
        Else
            frm.BackColor = -2147483633 'Normal color
            frm.TextBox2.BackColor = -2147483633 'Normal color
            frm.tBx1.BackColor = -2147483633 'Normal color
        End If
            frm.TextBox2.Value = ""
            frm.CommandButton2.caption = "Start"
            OngoingTimer = False
                
            'Redo basic calculations form the initialize macro
            M = Int(AllowedTime)
            S = (AllowedTime - Int(AllowedTime)) * 60
             With frm.tBx1
                .Value = Format(CStr(M), "00") & ":" & Format(CStr(S), "00")
            End With
    
            Call Optimize_VBA_Performance(False, xlAutomatic)
    End If
End Sub
