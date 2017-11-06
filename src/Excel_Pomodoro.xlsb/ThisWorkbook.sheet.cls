Private Sub Workbook_Open()
    If Range("Run_in_seperate_instance").Value = True Then
        If Reopen_decision = True Then Call OpenItSelfInAnotherInstance
    End If
    If Range("Shortcut").Value = True Then
        Call Shortcuts
    End If
End Sub

Sub Shortcuts()
    Application.OnKey "{F10}", "PomodoroSession"
End Sub