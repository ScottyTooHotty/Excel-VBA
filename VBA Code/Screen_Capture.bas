Attribute VB_Name = "Screen_Capture"
Option Explicit

Public Sub ScreenPrint()

'Was originally used with CDO mailer to give us an idea of what was happening at the point of error
'Currently unused

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    DoEvents
    Application.SendKeys "{1068}"
    
    Charts.Add
    ThisWorkbook.Charts(1).AutoScaling = True
    ThisWorkbook.Charts(1).Paste
    ThisWorkbook.Charts(1).Export Filename:="C:\Users\Public\Error - " & Format(Date, "yyyymmdd") & ".jpg", _
        FilterName:="jpg"
    
    ThisWorkbook.Charts(1).Delete
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub

