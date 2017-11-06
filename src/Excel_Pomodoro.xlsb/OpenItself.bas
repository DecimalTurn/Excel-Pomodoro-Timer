Attribute VB_Name = "OpenItself"
Option Explicit

Sub OpenItSelfInAnotherInstance()
    Dim objExcel As Excel.Application
    Set objExcel = CreateObject("Excel.Application")
    Dim FileName As String
    
    If VisibleWorkbookNB = 1 Then Workbooks.Add
    ShowWindow GetForegroundWindow, SW_SHOWMINIMIZED
    
    
    FileName = ThisWorkbook.FullName
         
    'ThisWorkbook.Save
    Dim TestString As String
    TestString = CStr(Application.Caller)
    If TestString = "Error 2023" Then
        ThisWorkbook.Saved = True
    Else
        ThisWorkbook.Save
    End If
    
    'Need to be in read-only mode before it opening itself in another instance of Excel
    ThisWorkbook.ChangeFileAccess xlReadOnly
     
    On Error GoTo Err1
    Call objExcel.Workbooks.Open(FileName)
    On Error GoTo 0
    
    objExcel.Visible = True
    objExcel.WindowState = xlMaximized
    
    ThisWorkbook.Close False
    
    Exit Sub
Err1:
    MsgBox "An error occured while trying to open the parrallel version of the file." & _
    "This could be due to the fact that the server hasn't release the file properly." & _
    "Save the file on your desktop might solve your problem."

    ThisWorkbook.Close False
End Sub

Public Function Reopen_decision() As Boolean


'Is there only one instance of Excel?
    Dim OnlyOne As Boolean
    If ExcelInstances = 1 Then OnlyOne = True

'Is it in the first instance
    Dim InFirst As Boolean
    'Get handle on the first instance
    Dim xlApp As Excel.Application
    Set xlApp = GetObject(, "Excel.Application")
    
    'Check if a workbook with thisworkbook name is open there.
    Dim wb As Workbook
    On Error Resume Next
    Set wb = xlApp.Application.Workbooks(ThisWorkbook.Name)
    On Error GoTo 0
    If Not wb Is Nothing Then
        InFirst = True
    Else
        InFirst = False
    End If
    
'Is the actual file both in another instance of Excel and the one in the first instance is just a copy?
    'Idea: let's compare the windows handle propertie to make sure they are different.
    Dim NotInFirstActually As Boolean
    
    If InFirst Then
        If xlApp.hWnd <> ThisWorkbook.Parent.hWnd Then NotInFirstActually = True
    End If
'Is the file alone?
    Dim Alone As Boolean
    If VisibleWorkbookNB = 1 Then Alone = True

'Has the file ever been saved?
    Dim FileEverSaved As Boolean
    If ThisWorkbook.Path <> "" Then
        FileEverSaved = True
    Else
        MsgBox "Warning: To work properly, the file needs to be saves somewhere on your computer.", vbCritical
    End If
    
'Create choice variable
    Dim i(1 To 5) As Integer
    'Convert our booleans into 1s and 0s
    i(1) = Abs(OnlyOne)
    i(2) = Abs(InFirst)
    i(3) = Abs(NotInFirstActually)
    i(4) = Abs(Alone)
    i(5) = Abs(FileEverSaved)

Dim choice_vr As String
    choice_vr = i(1) & i(2) & i(3) & i(4) & i(5)
    'Now that we have all the relevant information to treat our decision tree, we can proceed
    
    Dim DECISION As Boolean
    
    Select Case choice_vr 'See the open itself example file to view the decision tree
    Case Is = "00000": DECISION = 0
    Case Is = "00010": DECISION = 0
    Case Is = "00100": DECISION = 0
    Case Is = "00110": DECISION = 0
    Case Is = "01000": DECISION = 0
    Case Is = "01010": DECISION = 0
    Case Is = "01100": DECISION = 0
    Case Is = "01110": DECISION = 0
    Case Is = "10000": DECISION = 0
    Case Is = "10010": DECISION = 0
    Case Is = "10100": DECISION = 0
    Case Is = "10110": DECISION = 0
    Case Is = "11000": DECISION = 0
    Case Is = "11010": DECISION = 0
    Case Is = "11100": DECISION = 0
    Case Is = "11110": DECISION = 0

    Case Is = "00001": DECISION = 1
    Case Is = "00011": DECISION = 0
    Case Is = "00101": DECISION = 1
    Case Is = "00111": DECISION = 0
    Case Is = "01001": DECISION = 1
    Case Is = "01011": DECISION = 1
    Case Is = "01101": DECISION = 1
    Case Is = "01111": DECISION = 0
    Case Is = "10001": DECISION = 1
    Case Is = "10011": DECISION = 1
    Case Is = "10101": DECISION = 1
    Case Is = "10111": DECISION = 1
    Case Is = "11001": DECISION = 1
    Case Is = "11011": DECISION = 1
    Case Is = "11101": DECISION = 1
    Case Is = "11111": DECISION = 1

    Case Else: Err.Raise 13
    End Select

    Reopen_decision = DECISION

End Function


Function VisibleWorkbookNB()
Dim wb As Workbook, counter As Integer
    For Each wb In Excel.Application.Workbooks
        If wb.Windows(1).Visible = True Then
        counter = counter + 1
        End If
    Next wb

    VisibleWorkbookNB = counter
End Function
