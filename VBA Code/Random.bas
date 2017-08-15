Attribute VB_Name = "Random"
Option Explicit

Sub RandomRow()

Dim lAmount As Long
Dim lIterations As Long
Dim lRndRow As Long
Dim lFirstRow As Long
Dim lLastRow As Long
Dim vAmount As Variant
Dim vFirstRow As Variant
Dim wNewSht As Worksheet
Dim wSht As Worksheet
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    If ActiveWorkbook Is Nothing Then
        MsgBox "No open workbook to select random rows from." & vbNewLine & _
            vbNewLine & "Please open a workbook and try again", vbOKOnly Or _
            vbInformation, "No Open Workbook!"
        Exit Sub
    End If
    
    vFirstRow = _
        Application.InputBox("What row does your data start on?", "Staring Row")
        
    If vFirstRow = False Then
        GoTo Terminate
    End If
    
    vAmount = _
        Application.InputBox("How many rows do you want?", "Amount of Rows")
        
    If vAmount = False Then
        GoTo Terminate
    End If
    
    If Not IsNumeric(vAmount) Then
        GoTo InputError
    Else
        lAmount = vAmount
    End If
    
    Set wSht = Application.ActiveSheet

    lLastRow = wSht.Cells.Find(What:="*", After:=wSht.Cells(1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False).Row
    
    If Not IsNumeric(vFirstRow) Then
        GoTo InputError
    Else
        lFirstRow = vFirstRow
    End If
    
    ActiveWorkbook.Worksheets.Add After:=Sheets(Worksheets.Count)
    
    Set wNewSht = ActiveWorkbook.Worksheets(Worksheets.Count)
    
    If lFirstRow > 1 Then
        wSht.Range("1:" & lFirstRow - 1).Copy Destination:=wNewSht.Range("A1")
    End If
    
    lIterations = 1
    
    Do While lIterations <= lAmount
    
        lRndRow = Int((lLastRow - lFirstRow + 1) * Rnd + lFirstRow)
        
        wSht.Range(lRndRow & ":" & lRndRow).Copy _
            Destination:=wNewSht.Range("A" & lFirstRow + lIterations - 1)
        
        Debug.Print lRndRow
    
        lIterations = lIterations + 1
    
    Loop
    
    GoTo Terminate

InputError:
    MsgBox "You have not entered a valid number. Please try again", vbOKOnly Or _
        vbInformation, "Invalid entry"
    Exit Sub
    
Terminate:
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Sub
