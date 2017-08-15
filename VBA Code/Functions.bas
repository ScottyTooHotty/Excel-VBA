Attribute VB_Name = "Functions"
Option Explicit

Function FindColumn(strSearch As String, Optional ws As Worksheet, Optional rSearch As Range) As Long
'Searches for a specific string and returns the column number when found
'If ws is omitted then activesheet will be used
'If rSearch is omitted then whole worksheet will be used
'Function returns 0 if column not found

Dim rColumn As Range

    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If rSearch Is Nothing Then
        Set rSearch = Cells
    End If
    
    With ws.Range(rSearch.Address)
        Set rColumn = .Find(What:=strSearch, After:=rSearch.Cells(rSearch.Rows.Count, rSearch.Columns.Count), _
            LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, MatchCase:=False)
    End With
    
    If Not rColumn Is Nothing Then
        FindColumn = rColumn.Column
    Else
        FindColumn = 0
    End If

End Function

Function FindRow(strSearch As String, Optional ws As Worksheet, Optional rSearch As Range) As Long
'Searches for a specific string and returns the row number when found
'If ws is omitted then activesheet will be used
'If rSearch is omitted then whole worksheet will be used
'Function returns 0 if row not found

Dim rRow As Range

    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If rSearch Is Nothing Then
        Set rSearch = Cells
    End If

    With ws.Range(rSearch.Address)
        Set rRow = .Find(What:=strSearch, After:=rSearch.Cells(rSearch.Rows.Count, rSearch.Columns.Count), _
            LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False)
    End With
    
    If Not rRow Is Nothing Then
        FindRow = rRow.Row
    Else
        FindRow = 0
    End If

End Function

Function FirstColumn(Optional strSearch As String, Optional ws As Worksheet, Optional rSearch As Range) As Long
'Searches for the first column in the given worksheet/range and returns the column number
'If strSearch is omitted then the wildcard will be used
'If ws is omitted then activesheet will be used
'If rSearch is omitted then whole worksheet will be used
'Function returns 0 if column not found

Dim iLookAt As Integer
Dim rFirstColumn As Range

    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If rSearch Is Nothing Then
        Set rSearch = Cells
    End If

    If strSearch = "" Then
        strSearch = "*"
        iLookAt = 2
    Else
        iLookAt = 1
    End If

    With ws.Range(rSearch.Address)
        Set rFirstColumn = .Find(What:=strSearch, After:=rSearch.Cells(rSearch.Rows.Count, rSearch.Columns.Count), _
            LookIn:=xlFormulas, LookAt:=iLookAt, SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, MatchCase:=False)
    End With
    
    If Not rFirstColumn Is Nothing Then
        FirstColumn = rFirstColumn.Column
    Else
        FirstColumn = 0
    End If

End Function

Function FirstRow(Optional strSearch As String, Optional ws As Worksheet, Optional rSearch As Range) As Long
'Searches for the first row in the given worksheet/range and returns the row number
'If strSearch is omitted then the wildcard will be used
'If ws is omitted then activesheet will be used
'If rSearch is omitted then whole worksheet will be used
'Function returns 0 if row not found

Dim iLookAt As Integer
Dim rFirstRow As Range

    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If rSearch Is Nothing Then
        Set rSearch = Cells
    End If

    If strSearch = "" Then
        strSearch = "*"
        iLookAt = 2
    Else
        iLookAt = 1
    End If

    With ws.Range(rSearch.Address)
        Set rFirstRow = .Find(What:=strSearch, After:=rSearch.Cells(rSearch.Rows.Count, rSearch.Columns.Count), _
            LookIn:=xlFormulas, LookAt:=iLookAt, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False)
    End With
    
    If Not rFirstRow Is Nothing Then
        FirstRow = rFirstRow.Row
    Else
        FirstRow = 0
    End If

End Function

Function LastColumn(Optional ws As Worksheet, Optional rSearch As Range) As Long
'Searches for the last column in the given worksheet/range and returns the column number
'If ws is omitted then activesheet will be used
'If rSearch is omitted then whole worksheet will be used
'Function returns 0 if column not found

Dim rLastColumn As Range

    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If rSearch Is Nothing Then
        Set rSearch = Cells
    End If

    With ws.Range(rSearch.Address)
        Set rLastColumn = .Find(What:="*", After:=rSearch.Cells(1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
            MatchCase:=False)
    End With
    
    If Not rLastColumn Is Nothing Then
        LastColumn = rLastColumn.Column
    Else
        LastColumn = 0
    End If

End Function

Function LastRow(Optional ws As Worksheet, Optional rSearch As Range) As Long
'Searches for the last row in the given worksheet/range and returns the row number
'If ws is omitted then activesheet will be used
'If rSearch is omitted then whole worksheet will be used
'Function returns 0 if row not found

Dim rLastRow As Range

    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If rSearch Is Nothing Then
        Set rSearch = Cells
    End If

    With ws.Range(rSearch.Address)
        Set rLastRow = .Find(What:="*", After:=rSearch.Cells(1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
            MatchCase:=False)
    End With
    
    If Not rLastRow Is Nothing Then
        LastRow = rLastRow.Row
    Else
        LastRow = 0
    End If

End Function

Public Function SetRange(Optional strSearch As String, Optional ws As Worksheet, Optional rSearch As Range, _
    Optional lRowOffset As Long, Optional lColumnOffset As Long) As Range

Dim iLookAt As Integer
Dim rMyRange As Range
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If rSearch Is Nothing Then
        Set rSearch = Cells
    End If

    If strSearch = "" Then
        strSearch = "*"
        iLookAt = 2 '2 is xlPart
    Else
        iLookAt = 1 '1 is xlWhole
    End If
    
    With ws.Range(rSearch.Address)
        Set SetRange = .Find(What:=strSearch, After:=rSearch.Cells(rSearch.Rows.Count, rSearch.Columns.Count), _
            LookIn:=xlFormulas, LookAt:=iLookAt, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False)
    End With
    
    If Not SetRange Is Nothing Then
        Set SetRange = SetRange.Offset(lRowOffset, lColumnOffset)
    Else
        Debug.Print "Could not find " & strSearch
    End If

End Function
