Attribute VB_Name = "Reverse"
Option Explicit

Sub Reverse_Cell_Contents()

' Keyboard shortcut: Ctrl+r

Dim iLength As Integer
Dim sNew As String
Dim sRaw As String

   If Not ActiveCell.HasFormula Then
        sRaw = ActiveCell.Text
        sNew = ""
        For iLength = 1 To Len(sRaw)
            sNew = Mid(sRaw, iLength, 1) + sNew
        Next iLength
        ActiveCell.Value = sNew
   End If
End Sub



