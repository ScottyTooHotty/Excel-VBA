Attribute VB_Name = "Unprotect"
Option Explicit

Sub Password_Unprotect()
'Password for OneNote Workstreams and Project Planning Toolkit = "password"
'Password for UWMT VBA Project = "password"
'Password for UWMT Archive = "password"
'Password for Weekly Archive = "password" - Not protected yet

Const strPWord1 As String = "password"
Const strPWord2 As String = "password"
Const strPWord3 As String = "password"
Const strPWord4 As String = "password"
Const strPWord5 As String = "password"
Const strPWord6 As String = "password"
Const strPWord7 As String = "password"
    
    On Error Resume Next
        With ActiveSheet
            .Unprotect strPWord1
            .Unprotect strPWord2
            .Unprotect strPWord3
            .Unprotect strPWord4
            .Unprotect strPWord5
            .Unprotect strPWord6
            .Unprotect strPWord7
            
            .Unprotect
        End With
    On Error GoTo 0

End Sub
