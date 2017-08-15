Attribute VB_Name = "INV07Blanks"
Option Explicit

Public Sub BlankIDs()

Dim lChildIDCol(1 To 2) As Long
Dim lFirstRow(1 To 2) As Long
Dim lHeaderRow(1 To 2) As Long
Dim lINVIDCol(1 To 2) As Long
Dim lLastRow(1 To 2) As Long
Dim lOpenDateCol As Long
Dim rCell As Range
Dim rFind As Range
Dim rFound As Range
Dim rMax As Range
Dim wb As Workbook
Dim wbINV01 As Workbook
Dim wbINV07 As Workbook

    For Each wb In Application.Workbooks
        If wb.Name Like "*INV01*" Then
            Set wbINV01 = wb
        ElseIf wb.Name Like "*INV07*" Then
            Set wbINV07 = wb
        End If
    Next wb
    
    If wbINV01 Is Nothing Then
        MsgBox "INV01 cannot be found." & vbNewLine & "Make sure the workbook is open and try again.", _
            vbOKOnly Or vbInformation, "INV01"
        Exit Sub
    ElseIf wbINV07 Is Nothing Then
        MsgBox "INV07 cannot be found." & vbNewLine & "Make sure the workbook is open and try again.", _
            vbOKOnly Or vbInformation, "INV07"
        Exit Sub
    End If
    
    lChildIDCol(1) = FindColumn("Child Id", wbINV01.Sheets(1))
    lChildIDCol(2) = FindColumn("Child Id", wbINV07.Sheets(1))
    lINVIDCol(1) = FindColumn("Inv ID", wbINV01.Sheets(1))
    lINVIDCol(2) = FindColumn("Inv Id", wbINV07.Sheets(1))
    lOpenDateCol = FindColumn("Open Date", wbINV01.Sheets(1))
    lHeaderRow(1) = FindRow("Child Id", wbINV01.Sheets(1))
    lHeaderRow(2) = FindRow("Child Id", wbINV07.Sheets(1))
    lFirstRow(1) = lHeaderRow(1) + 1
    lFirstRow(2) = lHeaderRow(2) + 1
    lLastRow(1) = LastRow(wbINV01.Sheets(1))
    lLastRow(2) = LastRow(wbINV07.Sheets(1))
    
    With wbINV07.Sheets(1)
        For Each rCell In .Range(.Cells(lFirstRow(2), lChildIDCol(2)), .Cells(lLastRow(2), lChildIDCol(2)))
            If rCell.Value = "" Then
                Set rFound = FindAll(wbINV01.Sheets(1).Range(wbINV01.Sheets(1).Cells(lFirstRow(1), lINVIDCol(1)), _
                    wbINV01.Sheets(1).Cells(lLastRow(1), lINVIDCol(1))), .Cells(rCell.Row, lINVIDCol(2)).Value, _
                    xlFormulas, xlWhole, xlByRows, False)
                If Not rFound Is Nothing Then
                    Set rMax = Nothing
                    For Each rFind In rFound
                        If rMax Is Nothing Then
                            Set rMax = wbINV01.Sheets(1).Cells(rFind.Row, lOpenDateCol)
                        Else
                            If rMax.Value < wbINV01.Sheets(1).Cells(rFind.Row, lOpenDateCol) Then
                                Set rMax = wbINV01.Sheets(1).Cells(rFind.Row, lOpenDateCol)
                            End If
                        End If
                    Next rFind
                End If
                rCell.Value = wbINV01.Sheets(1).Cells(rMax.Row, lChildIDCol(1)).Value
            End If
        Next rCell
    End With
    
End Sub
