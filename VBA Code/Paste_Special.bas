Attribute VB_Name = "Paste_Special"
Option Explicit

Sub PasteSpecial()

    'PasteSpecialForm.Show
    
    With Application
        If .CutCopyMode = False Then
            Selection.Copy
        End If
        
        .Selection.PasteSpecial Paste:=xlPasteColumnWidths
        .Selection.PasteSpecial Paste:=xlPasteValues
        .Selection.PasteSpecial Paste:=xlPasteFormats
    
        .CutCopyMode = False
    End With

End Sub

