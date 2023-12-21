Option Explicit
Private Sub Workbook_Open()
    If Sheets("Settings").Range("Shortcut").Value = True And Not IsMac Then
        Shortcuts
    End If
    MacOptions
End Sub

Public Sub Shortcuts()
    Application.OnKey "{F10}", "PomodoroSession"
End Sub

Public Sub MacOptions()
'PURPOSE: Hide options that are not available for Mac

    Dim ss As Worksheet
    
    With ThisWorkbook.Sheets("Settings")
        If IsMac Then
            .Range("Reopen_Excel_after_x").EntireRow.Hidden = True
            .Range("Run_in_seperate_instance").EntireRow.Hidden = True
            .Range("Custom_position").EntireRow.Hidden = True
            .Range("Left_pos").EntireRow.Hidden = True
            .Range("Top_pos").EntireRow.Hidden = True
            .Range("Shortcut").EntireRow.Hidden = True
        Else
            .Range("Reopen_Excel_after_x").EntireRow.Hidden = False
            .Range("Run_in_seperate_instance").EntireRow.Hidden = False
            .Range("Custom_position").EntireRow.Hidden = False
            .Range("Left_pos").EntireRow.Hidden = False
            .Range("Top_pos").EntireRow.Hidden = False
            .Range("Shortcut").EntireRow.Hidden = False
        End If
    End With

End Sub