Attribute VB_Name = "Age"
Option Explicit

Public Function iAge(rDoB As Range, Optional dDate As Date) As Integer

    If dDate = 0 Then
        dDate = Date
    End If

    iAge = DateDiff("yyyy", rDoB, dDate) + (dDate < DateSerial(Year(dDate), Month(rDoB), Day(rDoB)))

End Function
