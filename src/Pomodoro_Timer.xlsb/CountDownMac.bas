Attribute VB_Name = "CountDownMac"
Option Explicit

Const FREQ As Double = 1

Public Sub LaunchTimerMac(ByVal Timer As PomodoroTimer)
    'Stop the code if the form is not visible
    If UFIsVisible = False Then: Debug.Print "Form is not visible. The code will now stop.": End
    
    Dim ThisForm As PomodoroTimer
    Set ThisForm = Timer
    OptimizeVbaPerformance True
    
    OngoingTimer = True
    StopTimer = False
    CloseTimer = False
    ThisForm.CommandButton2.caption = "Cancel"
    
    'Reset the colors
    PomodoroTimer.BackColor = -2147483633
    ThisForm.TextBox2.BackColor = -2147483633
    ThisForm.tBx1.BackColor = -2147483633
    
    StartTime = Now()
    TodaysDate = Date
    
    Dim SecondsRemaining As Double
    Dim MinutesRemaining As Double
    Dim TotalTime As Double
    Dim EndTime As Double
    Dim RemaingTime As Double
    
    TotalTime = 60 * AllowedTime + AllowedTimeSec
    EndTime = DateAdd("s", TotalTime, Now())
    RemaingTime = DateDiff("s", Now(), EndTime)
    
        RemaingTime = DateDiff("s", Now(), EndTime)
        MinutesRemaining = Int(RemaingTime / 60)
        SecondsRemaining = RemaingTime - 60 * MinutesRemaining
        
        With ThisForm.tBx1
            .Value = Format$(CStr(MinutesRemaining), "00") & ":" & Format$(CStr(SecondsRemaining), "00")
        End With
        
        'Released the control to the OS
        'DoEvents
        
        'Now "sleep"
        Application.OnTime Now + TimeValue("00:00:01") * FREQ, "LaunchTimerMac2"

End Sub

'@Ignore UseMeaningfulName
Public Sub LaunchTimerMac2(ByVal Timer As PomodoroTimer)
    
    Dim ThisForm As PomodoroTimer
    Set ThisForm = Timer
    
    '@Ignore UseMeaningfulName
    Dim M As Double, S As Double
    Dim TotalTime As Double
    Dim EllapsedtTime As Double
    Dim StartTime As Double
    Dim EndTime As Double
    Dim RemaingTime As Double
    
    TotalTime = 60 * AllowedTime + AllowedTimeSec
'    EllapsedtTime = TotalTime - (60 * Split(frm.tBx1.Value, ":")(0) + 1 * Split(frm.tBx1.Value, ":")(1))
'        M = Int(EllapsedtTime / 60)
'        S = EllapsedtTime - 60 * M
'    EndTime = DateAdd("s", TotalTime, Now())
'    StartTime = Now() - TimeValue("00:" & Format$(CStr(M), "00") & ":" & Format$(CStr(S), "00"))
    RemaingTime = 60 * Split(ThisForm.tBx1.Value, ":")(0) + 1 * Split(ThisForm.tBx1.Value, ":")(1)
    
    If RemaingTime > 0 And Not StopTimer Then
        RemaingTime = RemaingTime - FREQ
        M = Int(RemaingTime / 60)
        S = RemaingTime - 60 * M
        
        With ThisForm.tBx1
            .Value = Format$(CStr(M), "00") & ":" & Format$(CStr(S), "00")
        End With
        
        'Released the control to the OS
        'DoEvents
        
        'Now "sleep"
        Application.OnTime Now + TimeValue("00:00:01") * FREQ, "LaunchTimerMac2"
        
    Else
    
        'Since we are using the "Application.ontime" technique, it is possible that some public variables will have lost their values
        If TodaysDate = 0 Then TodaysDate = Date
        If StartTime = 0 Then
            EllapsedtTime = TotalTime - RemaingTime
            M = Int(EllapsedtTime / 60)
            S = EllapsedtTime - 60 * M
            StartTime = Now() - TimeValue("00:" & Format$(CStr(M), "00") & ":" & Format$(CStr(S), "00"))
        End If
        
        'Recording session
        If StopTimer = False Or ThisWorkbook.Sheets("Settings").Range("Record_unfinished").Value2 = True Then
            If (TotalTime - RemaingTime) / 60 > ThisWorkbook.Sheets("Settings").Range("No_Recording_limit") Then
                Add_new_record TodaysDate, StartTime, Now, Not (StopTimer), ThisWorkbook.Sheets("Pomodoro").Range("TaskNameRng")
            End If
        End If
        
        OptimizeVbaPerformance False, xlAutomatic
        
        If StopTimer = False Then 'If the timer was stopped by the user
            'Proceed with the Break
            If ThisWorkbook.Sheets("Settings").Range("Sound_end_Pomodoro") = True Then Beep
            ThisForm.TextBox2.Value = "Break"
            TakeBreakMac ThisForm
        Else
            'Do nothing
            ThisForm.CommandButton2.caption = "Start"
            OngoingTimer = False
        End If
        
        If CloseTimer Then Unload ThisForm
    End If
End Sub

Private Sub TakeBreakMac(ByVal Timer As PomodoroTimer)

    Dim ThisForm As PomodoroTimer
    Set ThisForm = Timer
    'Reset StopTimer:
    StopTimer = False
    
    OptimizeVbaPerformance True
    
    '@Ignore UseMeaningfulName
    '@Ignore UseMeaningfulName
    Dim M As Double, S As Double
    M = BreakTime
    S = BreakTimeSec
     
    With ThisForm.tBx1
        .Value = Format$(CStr(M), "00") & ":" & Format$(CStr(S), "00")
    End With

    TakeBreakMac2 ThisForm

End Sub

'@Ignore UseMeaningfulName
Private Sub TakeBreakMac2(ByVal Timer As PomodoroTimer)
    
    
    Dim frm As PomodoroTimer
    Set frm = Timer

    Dim SecondsRemaining As Long
    Dim MinutesRemaining As Long
    Dim EndTime As Double
    Dim RemaingTime As Double
    Dim TotalTime As Long
    
    TotalTime = 60 * BreakTime + BreakTimeSec
    RemaingTime = 60 * Split(frm.tBx1.Value, ":")(0) + 1 * Split(frm.tBx1.Value, ":")(1)
    
   If RemaingTime > 0 And Not StopTimer Then
        RemaingTime = RemaingTime - FREQ
        MinutesRemaining = Int(RemaingTime / 60)
        SecondsRemaining = RemaingTime - 60 * MinutesRemaining
        
        'Flashing
        If TotalTime - RemaingTime < 9 Then
            If SecondsRemaining Mod 2 = 1 Then
                frm.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
                frm.TextBox2.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
                frm.tBx1.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
            Else
                frm.BackColor = -2147483633 'Normal color
                frm.TextBox2.BackColor = -2147483633 'Normal color
                frm.tBx1.BackColor = -2147483633 'Normal color
            End If
        End If
        
        With frm.tBx1
            .Value = Format$(CStr(MinutesRemaining), "00") & ":" & Format$(CStr(SecondsRemaining), "00")
        End With
        'Released the control to the OS
        'DoEvents
        'Now "sleep"
        Application.OnTime Now + TimeValue("00:00:01") * FREQ, "TakeBreakMac2"
        
    Else

        If StopTimer = False Then
            If ThisWorkbook.Sheets("Settings").Range("Sound_end_Break") = True Then Beep
            'Remain in color to get the user's attention
            frm.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
            frm.TextBox2.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
            frm.tBx1.BackColor = GetRGBColor_Fill(ThisWorkbook.Sheets("Settings").Range("Flashing_color")) 'Flashing color
        Else
            frm.BackColor = -2147483633 'Normal color
            frm.TextBox2.BackColor = -2147483633 'Normal color
            frm.tBx1.BackColor = -2147483633 'Normal color
        End If
            frm.TextBox2.Value = vbNullString
            frm.CommandButton2.caption = "Start"
            OngoingTimer = False
                
            'Redo basic calculations form the initialize macro
            MinutesRemaining = Int(AllowedTime)
            SecondsRemaining = (AllowedTime - Int(AllowedTime)) * 60
             With frm.tBx1
                .Value = Format$(CStr(MinutesRemaining), "00") & ":" & Format$(CStr(SecondsRemaining), "00")
            End With
    
            OptimizeVbaPerformance False, xlAutomatic
    End If
End Sub
