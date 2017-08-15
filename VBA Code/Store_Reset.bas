Attribute VB_Name = "Store_Reset"
Option Explicit

Sub Store_Workbook_State_v2()

Dim iCount As Integer, lCurrentRow As Long, iLastColumn As Integer
Dim iStringLength As Integer, iVisible As Integer
Dim lLastRow As Long
Dim sFirstCell As String, sFirstSheet As String, sNew As String, sPassword As String
Dim sRaw As String, sSht As String
Dim wWs As Worksheet

    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .DisplayStatusBar = True
    End With
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "No open workbook to save. Please open a workbook and try again", _
            vbOKOnly Or vbInformation, "No Open Workbook!"
        GoTo Terminate
    End If

    For Each wWs In ActiveWorkbook.Worksheets
        If wWs.Name = "Workbook State Macro" Then
            sSht = wWs.Name
            Exit For
        End If
    Next wWs
    
    sFirstSheet = ActiveWorkbook.ActiveSheet.Name
        
    If sSht = "" Then
        With ActiveWorkbook
            .Sheets.Add
            .ActiveSheet.Name = "Workbook State Macro"
            sSht = .Sheets("Workbook State Macro").Name
        End With
    Else
        sRaw = Sheets(sSht).Range("A2").Value
        sNew = ""
        For iStringLength = 1 To Len(sRaw)
            sNew = Mid(sRaw, iStringLength, 1) + sNew
        Next iStringLength
        sPassword = sNew
        ActiveWorkbook.Sheets(sSht).Unprotect Password:=sPassword
    End If
    
    With ActiveWorkbook.Sheets(sSht)
    
        .Cells.Clear
        .Range("A1").Value = "Sheet Name"
        .Range("B1").Value = "Last Row"
        .Range("C1").Value = "Last Column"
        .Range("D1").Value = "Named Range"
        lCurrentRow = 2

        For Each wWs In ActiveWorkbook.Worksheets
            
            If wWs.Name = "Workbook State Macro" Then
                GoTo NextwWs
            End If
            
            Application.StatusBar = "Saving state of " & wWs.Name
            
            If wWs.Visible <> xlSheetVisible Then
                iVisible = wWs.Visible
                wWs.Visible = xlSheetVisible
            Else
                iVisible = wWs.Visible
            End If
            
            wWs.Select
            sFirstCell = ActiveCell.Address
            
            On Error GoTo Err1
            
            lLastRow = _
                wWs.Cells.Find(What:="*", After:=wWs.Cells(1), LookAt:=xlPart, _
                LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, MatchCase:=False).Row

            iLastColumn = _
                wWs.Cells.Find(What:="*", After:=wWs.Cells(1), LookAt:=xlPart, _
                LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, MatchCase:=False).Column

            .Cells(lCurrentRow, 1).Value = wWs.Name

            If lLastRow = 1 And iLastColumn = 1 Then
                .Cells(lCurrentRow, 2).Value = lLastRow
                .Cells(lCurrentRow, 3).Value = iLastColumn
            Else
                .Cells(lCurrentRow, 2).Value = lLastRow + 1
                .Cells(lCurrentRow, 3).Value = iLastColumn + 1
            End If
            
            On Error GoTo Err2
            
            '.Cells.SpecialCells(xlCellTypeBlanks).Select
            
            ActiveWorkbook.Names.Add Name:=wWs.CodeName & "Blank", _
                RefersToR1C1:=wWs.Cells.SpecialCells(xlCellTypeBlanks)
            
            On Error GoTo 0
            
            .Cells(lCurrentRow, 4).Value = wWs.CodeName & "Blank"
            
            lCurrentRow = lCurrentRow + 1
            
            wWs.Range(sFirstCell).Select
            
            If iVisible <> xlVisible Then
                wWs.Visible = iVisible
            End If
NextwWs:
        Next wWs
        
        sRaw = Sheets(sSht).Range("A2").Value
        sNew = ""
        For iStringLength = 1 To Len(sRaw)
            sNew = Mid(sRaw, iStringLength, 1) + sNew
        Next iStringLength
        sPassword = sNew
        
        .Protect Password:=sPassword, DrawingObjects:=True, Contents:=True, _
            Scenarios:=True
        .EnableSelection = xlNoSelection
        .Visible = xlVeryHidden
        
    End With
    
    If Err.Number = 0 Then
        MsgBox "The workbook state has been saved successfully", vbOKOnly Or _
            vbInformation, "Saved"
    Else
        GoTo Err2
    End If

    If iVisible <> -1 Then
        wWs.Visible = iVisible
    End If
    
    If sFirstSheet <> sSht Then
        Worksheets(sFirstSheet).Select
    End If
    
    GoTo Terminate
    
Err1:
    If Err.Number = 91 Then
        lLastRow = 1
        iLastColumn = 1
        Err.Clear
        Resume Next
    ElseIf Err.Number <> 0 Then
        MsgBox "Error number " & Err.Number & " encountered" & vbCrLf & _
            Err.Description, vbOKOnly Or vbCritical, "Error occurred"
        On Error Resume Next
        Sheets(sSht).Delete
        'Err.Clear
        GoTo Terminate
    End If
        'MsgBox "The workbook state has been saved successfully", vbOKOnly Or _
            vbInformation, "Saved"
        'Err.Clear
    'End If
    
Err2:
    If Err.Number = 1004 Then
        'wWs.Cells(Sheets(sSht).Cells(lCurrentRow, 2).Value, Sheets(sSht). _
            Cells(lCurrentRow, 3).Value).Select
        ActiveWorkbook.Names.Add Name:=wWs.CodeName & "Blank", _
            RefersToR1C1:=wWs.Cells(Sheets(sSht).Cells(lCurrentRow, 2).Value, _
            Sheets(sSht).Cells(lCurrentRow, 3).Value)
        MsgBox "The " & wWs.Name & " worksheet is either blank or the used range" & _
            " is too complex to save." & vbCrLf & vbCrLf & "Only the last row" & _
            " & column have been found so all data changed within the threshold" & _
            " of the" & vbCrLf & "last row & column will not reset correctly.", _
                vbOKOnly Or vbInformation, "Used range too complex"
        Err.Clear
        Resume Next
    Else
        If Err.Number <> 0 Then
            GoTo Err1
        End If
    End If
    
Terminate:
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .StatusBar = False
    End With
    
End Sub

Sub Reset_Workbook_State()

Dim lCurrentRow As Long, iTempLastColumn As Integer
Dim lLastRow As Long, lTempLastRow As Long
Dim sNamedRange As String, sSht As String, sTempSht As String
Dim wWs As Worksheet

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .DisplayStatusBar = True
    End With
    
    On Error GoTo ErrHnd
        
    If ActiveWorkbook Is Nothing Then
        MsgBox "No open workbook. Please open a workbook and try again", vbOKOnly _
            Or vbInformation, "No Open Workbook!"
        GoTo Terminate
    End If

    For Each wWs In ActiveWorkbook.Worksheets
        If wWs.Name = "Workbook State Macro" Then
            sSht = wWs.Name
        End If
    Next wWs
    
    If sSht = "" Then
        MsgBox "The workbook state has not been saved so recall is impossible", _
            vbOKOnly Or vbExclamation, "Unable to reset workbook"
        GoTo Terminate
    End If
    
    lLastRow = Sheets(sSht).Cells.Find(What:="*", After:=Sheets(sSht).Cells(1), _
        LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False).Row
    lCurrentRow = 2
    
    Do While lCurrentRow <> lLastRow + 1 And sTempSht <> sSht
        With ActiveWorkbook.Sheets(sSht)
            sTempSht = .Cells(lCurrentRow, 1).Value
            lTempLastRow = .Cells(lCurrentRow, 2).Value
            iTempLastColumn = .Cells(lCurrentRow, 3).Value
            sNamedRange = .Cells(lCurrentRow, 4).Value
            
            If ActiveWorkbook.Sheets(sTempSht) Is Nothing Then
                GoTo ErrHnd
            End If
            
            With ActiveWorkbook.Sheets(sTempSht)
            
                Application.StatusBar = "Restoring state of " & sTempSht
                
                With .Range(.Rows(lTempLastRow), .Rows(Rows.Count))
                    .ClearContents
                    .ClearOutline
                    .ClearNotes
                    .ClearComments
                End With
                With .Range(.Columns(iTempLastColumn), .Columns(Columns.Count))
                    .ClearContents
                    .ClearOutline
                    .ClearNotes
                    .ClearComments
                End With
                On Error Resume Next
                With .Range(sNamedRange)
                    .ClearContents
                    .ClearOutline
                    .ClearNotes
                    .ClearComments
                End With
                If Err.Number <> 0 Then
                    MsgBox "There is a problem with the named range '" & _
                        sNamedRange & "'." & vbCrLf & "It could be " & _
                        "missing, corrupted or contain merged cells, and so this " & _
                        "step has been skipped." & vbCrLf & "This may mean that " & _
                        "the sheet '" & sTempSht & "' has not been restored " & _
                        "properly." & vbCrLf & "Please check this sheet " & _
                        "thoroughly before assuming it has been reset.", _
                        vbOKOnly Or vbExclamation, "Named Range Issue"
                    Err.Clear
                End If
                On Error GoTo ErrHnd
            End With
        End With
    lCurrentRow = lCurrentRow + 1
    Loop
    
    MsgBox "The workbook state has been restored", vbOKOnly Or vbInformation, _
        "Restored"
        
    GoTo Terminate
                
ErrHnd:

    MsgBox "Error " & Err.Number & " encountered in Reset_Workbook_State." & _
        vbCrLf & Err.Description & vbCrLf & "Error resetting " & sTempSht, _
        vbOKOnly Or vbInformation, "Error"
                
Terminate:
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = False
    End With

End Sub


