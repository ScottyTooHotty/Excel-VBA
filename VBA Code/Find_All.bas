Attribute VB_Name = "Find_All"
Option Explicit

Public Function FindAll(rSearchRange As Range, vFindWhat As Variant, _
    Optional xlfLookIn As XlFindLookIn = xlValues, Optional xllLookAt As XlLookAt = xlWhole, _
    Optional xlsSearchOrder As XlSearchOrder = xlByRows, Optional bMatchCase As Boolean = False, _
    Optional strBeginsWith As String = vbNullString, _
    Optional strEndsWith As String = vbNullString, _
    Optional vbcBeginEndCompare As VbCompareMethod = vbTextCompare) As Range

Dim rFoundCell As Range
Dim rFirstFound As Range
Dim rLastCell As Range
Dim rResultRange As Range
Dim xllXLookAt As XlLookAt
Dim bInclude As Boolean
Dim vbcCompMode As VbCompareMethod
Dim rArea As Range
Dim lMaxRow As Long
Dim lMaxCol As Long
Dim bBegin As Boolean
Dim bEnd As Boolean


    vbcCompMode = vbcBeginEndCompare
    If strBeginsWith <> vbNullString Or strEndsWith <> vbNullString Then
        xllXLookAt = xlPart
    Else
        xllXLookAt = xllLookAt
    End If

    For Each rArea In rSearchRange.Areas
        With rArea
            If .Cells(.Cells.Count).Row > lMaxRow Then
                lMaxRow = .Cells(.Cells.Count).Row
            End If
            If .Cells(.Cells.Count).Column > lMaxCol Then
                lMaxCol = .Cells(.Cells.Count).Column
            End If
        End With
    Next rArea
    Set rLastCell = rSearchRange.Worksheet.Cells(lMaxRow, lMaxCol)

    On Error GoTo 0
    Set rFoundCell = rSearchRange.Find(What:=vFindWhat, After:=rLastCell, _
        LookIn:=xlfLookIn, LookAt:=xllXLookAt, SearchOrder:=xlsSearchOrder, _
        MatchCase:=bMatchCase)

    If Not rFoundCell Is Nothing Then
        Set rFirstFound = rFoundCell
        Do Until False
            bInclude = False
            If strBeginsWith = vbNullString And strEndsWith = vbNullString Then
                bInclude = True
            Else
                If strBeginsWith <> vbNullString Then
                    If StrComp(Left(rFoundCell.Text, Len(strBeginsWith)), strBeginsWith, _
                        vbcBeginEndCompare) = 0 Then
                        bInclude = True
                    End If
                End If
                If strEndsWith <> vbNullString Then
                    If StrComp(Right(rFoundCell.Text, Len(strEndsWith)), strEndsWith, _
                        vbcBeginEndCompare) = 0 Then
                        bInclude = True
                    End If
                End If
            End If
            If bInclude = True Then
                If rResultRange Is Nothing Then
                    Set rResultRange = rFoundCell
                Else
                    Set rResultRange = Application.Union(rResultRange, rFoundCell)
                End If
            End If
            Set rFoundCell = rSearchRange.FindNext(After:=rFoundCell)
            If (rFoundCell Is Nothing) Then
                Exit Do
            End If
            If (rFoundCell.Address = rFirstFound.Address) Then
                Exit Do
            End If
    
        Loop
    End If
    
    Set FindAll = rResultRange

End Function

