Attribute VB_Name = "Find_Username"
Option Explicit

Public Function UserName() As String

Dim iIndex As Integer
Dim oUser As Object
Dim strName As String
Dim strTemp As String
Dim strTempArray() As String
    
    On Error Resume Next
    strTemp = CreateObject("ADSystemInfo").UserName
    Set oUser = GetObject("LDAP://" & strTemp)

    If Err.Number = 0 Then
        strName = oUser.Get("givenName") & Chr(32) & oUser.Get("sn")
    Else
        strTempArray = Split(strTemp, ",")
        For iIndex = 0 To UBound(strTempArray)
            If UCase(Left(strTempArray(iIndex), 3)) = "CN=" Then
                strName = Trim(Mid(strTempArray(iIndex), 4))
                Exit For
            End If
        Next
    End If
    
    On Error GoTo 0
    
    UserName = strName
    
End Function


