Attribute VB_Name = "Remove_UsedRange"
Option Explicit

Sub Remove_UsedRange()

Dim iAnswer As Integer
Dim lLastColumn As Long
Dim lLastRow As Long
Dim lShapeColumn As Long
Dim lShapeRow As Long
Dim rColumnFormula As Range
Dim rColumnValue As Range
Dim rRowFormula As Range
Dim rRowValue As Range
Dim sShp As Shape
Dim wSht As Worksheet
   
    On Error GoTo ErrHnd
     
    For Each wSht In Worksheets
    
    If wSht.ProtectContents = True Then
        iAnswer = MsgBox("Unable to remove the used range on the sheet '" & _
            wSht.Name & "' as it is protected." & vbCrLf & "Click OK to skip " & _
            "this sheet and carry on, or Cancel to stop now." & vbCrLf & vbCrLf & _
            "To get the best results, unprotect all worksheets and run the " & _
            "procedure again.", vbOKCancel Or vbInformation, "Sheet Protected")
        If iAnswer = vbCancel Then
            Exit Sub
        Else
            GoTo NextwSht
        End If
    End If
    
        With wSht
    
            On Error Resume Next
            Set rColumnFormula = .Cells.Find(What:="*", After:=.Cells(1), _
                LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious)
            Set rColumnValue = .Cells.Find(What:="*", After:=.Cells(1), _
                LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious)
            Set rRowFormula = .Cells.Find(What:="*", After:=.Cells(1), _
                LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious)
            Set rRowValue = .Cells.Find(What:="*", After:=.Cells(1), _
                LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious)
            On Error GoTo 0
       
            If rRowFormula Is Nothing Then
                lLastRow = 0
            Else
                lLastRow = rRowFormula.Row
            End If
            If Not rRowValue Is Nothing Then
                lLastRow = Application.WorksheetFunction.Max(lLastRow, rRowValue.Row)
            End If
       
            If rColumnFormula Is Nothing Then
                lLastColumn = 0
            Else
                lLastColumn = rColumnFormula.Column
            End If
            If Not rColumnValue Is Nothing Then
                lLastColumn = Application.WorksheetFunction.Max(lLastColumn, _
                    rColumnValue.Column)
            End If

            For Each sShp In .Shapes
                lShapeRow = 0
                lShapeColumn = 0
                On Error Resume Next
                lShapeRow = sShp.TopLeftCell.Row
                lShapeColumn = sShp.TopLeftCell.Column
                On Error GoTo 0
                If lShapeRow > 0 And lShapeColumn > 0 Then
                    Do Until .Cells(lShapeRow, lShapeColumn).Top > _
                        sShp.Top + sShp.Height
                        lShapeRow = lShapeRow + 1
                    Loop
                If lShapeRow > lLastRow Then
                    lLastRow = lShapeRow
                End If
                Do Until .Cells(lShapeRow, lShapeColumn).Left > sShp.Left + sShp.Width
                    lShapeColumn = lShapeColumn + 1
                Loop
                    If lShapeColumn > lLastColumn Then
                        lLastColumn = lShapeColumn
                    End If
                End If
            Next
       
            .Range(.Cells(1, lLastColumn + 1).Address, .Cells(Rows.Count, _
                Columns.Count)).Delete
            .Range(.Cells(lLastRow + 1, 1).Address, .Cells(Rows.Count, _
                Columns.Count)).Delete
        End With
NextwSht:
    Next wSht
    
    iAnswer = MsgBox("Click OK to save the workbook, or cancel to skip." & vbCrLf & _
        "If you do not save now, the usedrange will remain until you do save.", _
        vbOKCancel Or vbExclamation, "Save Now?")
        
    If iAnswer = vbOK Then
        ActiveWorkbook.Save
    End If
  
Exit Sub

ErrHnd:

    MsgBox "Error " & Err.Number & " encountered in Remove_UsedRange" & vbCrLf & _
        Err.Description, vbOKOnly Or vbInformation, "Error"
  
End Sub
