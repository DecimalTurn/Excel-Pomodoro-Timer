Attribute VB_Name = "API_Sleep"
'PURPOSE: Define the sleep function to stop the code from running and releasing CPU usage.

Option Explicit

#If Mac Then
    #If MAC_OFFICE_VERSION >= 15 Then
        #If VBA7 Then ' 64-bit Excel 2016 for Mac
            
            Public Declare PtrSafe Sub Sleep _
            Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" _
            (ByVal dwMilliseconds As Long)
        #Else ' 32-bit Excel 2016 for Mac
        
            Public Declare Sub Sleep _
            Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" _
            (ByVal dwMilliseconds As Long)
            
        #End If
    #Else ' 32-bit Excel 2011 for Mac
            Public Declare Sub Sleep _
            Lib "Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" _
            (ByVal dwMilliseconds As Long)
    #End If
#Else
    #If VBA7 Then ' Excel 2010 or later for Windows
    
        Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 64 Bit Systems
    
    #Else ' pre Excel 2010 for Windows
        
        Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
    
    #End If
#End If


Sub SleepTest()
'MsgBox "Execution is started"
Sleep 10000 'delay in milliseconds
MsgBox "Waiting completed"
End Sub
