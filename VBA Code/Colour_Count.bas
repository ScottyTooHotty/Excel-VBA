Attribute VB_Name = "Colour_Count"
Option Explicit

Function ColourCount(rMyRange As Range, Optional rOriginal As Range) As Long

Dim iColour As Integer
Dim lCount As Long
Dim rCell As Range
    
    Application.Volatile True
    
    If rOriginal Is Nothing Then
        Set rOriginal = ActiveSheet.Range(Application.Caller.Address)
    End If
    
    lCount = 0

    For Each rCell In rMyRange
        If rCell.Interior.colorindex <> xlNone Then
            If rCell.Interior.colorindex = rOriginal.Interior.colorindex Then
                lCount = lCount + 1
            End If
        End If
    Next rCell
    
    ColourCount = lCount

End Function
