Attribute VB_Name = "ColourFunctions"
Option Explicit

Public Function FontColour(Optional rCell As Range) As Variant

Dim vColour As Variant

    Application.Volatile

    If rCell Is Nothing Then
        Set rCell = Application.Caller
    End If

    If rCell.Cells.Count > 1 Then
        vColour = rCell(1).Font.Color
    Else
        vColour = rCell.Font.Color
    End If
    
    If vColour = -4105 Then
        vColour = "Automatic"
    End If
    
    FontColour = vColour

End Function

Public Function FontColourIndex(Optional rCell As Range) As Variant

Dim vColour As Variant
    
    Application.Volatile

    If rCell Is Nothing Then
        Set rCell = Application.Caller
    End If

    If rCell.Cells.Count > 1 Then
        vColour = rCell(1).Font.colorindex
    Else
        vColour = rCell.Font.colorindex
    End If
    
    If vColour = -4105 Then
        vColour = "Automatic"
    End If
    
    FontColourIndex = vColour

End Function

Public Function FillColour(Optional rCell As Range) As Variant

Dim vColour As Variant

    Application.Volatile

    If rCell Is Nothing Then
        Set rCell = Application.Caller
    End If

    If rCell.Cells.Count > 1 Then
        vColour = rCell(1).Interior.Color
    Else
        vColour = rCell.Interior.Color
    End If
    
    If vColour = -4142 Then
        vColour = "No Fill"
    End If
    
    FillColour = vColour

End Function

Public Function FillColourIndex(Optional rCell As Range) As Variant

Dim vColour As Variant

    Application.Volatile

    If rCell Is Nothing Then
        Set rCell = Application.Caller
    End If

    If rCell.Cells.Count > 1 Then
        vColour = rCell(1).Interior.colorindex
    Else
        vColour = rCell.Interior.colorindex
    End If
    
    If vColour = -4142 Then
        vColour = "No Fill"
    End If
    
    FillColourIndex = vColour

End Function
