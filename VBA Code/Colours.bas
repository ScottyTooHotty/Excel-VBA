Attribute VB_Name = "Colours"
Option Explicit

Sub Fill_Colour()
Attribute Fill_Colour.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Keyboard Shortcut: Ctrl+q

Dim rCell As Range
Dim ColourArray As Variant
Dim Colour1 As Variant
Dim Mtch As Variant

    If Application.VERSION <= 13 Then
        ColourArray = Array(3, 35, 36, 37, 15, xlNone)
        If Selection.Cells.Count > 1 Then
            Colour1 = Selection.SpecialCells(xlCellTypeVisible).Interior.colorindex
        Else
            Colour1 = Selection.Interior.colorindex
        End If
        Mtch = Application.Match(Colour1, ColourArray, 0)
        'On Error GoTo 0
        If IsError(Mtch) Then
            Mtch = 6
        Else
            Mtch = Mtch + 1
            If Mtch = 7 Then Mtch = 1
        End If
        Selection.Interior.colorindex = Application.Index(ColourArray, Mtch)
    Else
        ColourArray = Array(255, 13434828, 10092543, 16764057, 12632256, 16777215)
        If Selection.Cells.Count > 1 Then
            Colour1 = Selection.SpecialCells(xlCellTypeVisible).Interior.Color
        Else
            Colour1 = Selection.Interior.Color
        End If
        Mtch = Application.Match(Colour1, ColourArray, 0)
        'On Error GoTo 0
        If IsError(Mtch) Then
            Mtch = 6
        Else
            Mtch = Mtch + 1
            If Mtch = 7 Then Mtch = 1
        End If
        If Application.Index(ColourArray, Mtch) = 16777215 Then
            Selection.Interior.colorindex = xlNone
        Else
            Selection.Interior.Color = Application.Index(ColourArray, Mtch)
        End If
    End If
    
End Sub

Sub Font_Colour()
Attribute Font_Colour.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' Keyboard Shortcut: Ctrl+w

Dim ColourArray As Variant
Dim Colour1 As Variant
Dim Mtch As Variant

    ColourArray = Array(3, 4, 41, 2, 0)
    Colour1 = Selection.Font.colorindex
    Mtch = Application.Match(Colour1, ColourArray, 0)
   ' On Error GoTo 0
    If IsError(Mtch) Then
        Mtch = 1
    Else
        Mtch = Mtch + 1
        If Mtch = 6 Then Mtch = 1
    End If
    Selection.Font.colorindex = Application.Index(ColourArray, Mtch)
    
End Sub

