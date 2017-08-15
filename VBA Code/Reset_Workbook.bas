Attribute VB_Name = "Reset_Workbook"
Option Explicit

Sub Reset_Workbook()

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .CellDragAndDrop = True
        .StatusBar = False
    End With

End Sub

