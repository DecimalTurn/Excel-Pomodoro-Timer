Attribute VB_Name = "vbaOptimizeRuntime"
Option Explicit

Sub OptimizeVbaPerformance(ByVal Optimize As Boolean, Optional Calculation As Variant)
'PURPOSE: Disable some VBA related events to allow the code to run faster
'MORE INFOS: http://analystcave.com/excel-improve-vba-performance/
'https://support.microsoft.com/fr-fr/help/199505/macro-performance-slow-when-page-breaks-are-visible-in-excel
    If IsMissing(Calculation) Then
        Calculation = IIf(Optimize, xlCalculationManual, xlCalculationAutomatic)
    End If
    With Application
        .Calculation = Calculation
        .ScreenUpdating = Not (Optimize)
        .EnableEvents = Not (Optimize)
    End With
End Sub
