Attribute VB_Name = "Calendar"
Option Explicit

Public dFinalDate As Date

Sub DatePicker()

    DatePickerForm.Show

End Sub

Sub GenerateCalendar(dCurrent As Date)

Dim ctrlButton As Control
Dim dDate As Date
Dim iCount As Integer
Dim iDay As Integer
Dim iDayofWeek As Integer
Dim iMonth As Integer
Dim iYear As Integer
Dim strDayButton As String

    iMonth = Month(dCurrent)
    iDay = 1
    iYear = Year(dCurrent)
    dDate = DateSerial(iYear, iMonth, iDay)
    
    iDayofWeek = Weekday(dDate, vbMonday)
    
    Select Case iDayofWeek
        'Monday
        Case 1
            iCount = 1
        'Tuesday
        Case 2
            iCount = 2
        'Wednesday
        Case 3
            iCount = 3
        'Thursday
        Case 4
            iCount = 4
        'Friday
        Case 5
            iCount = 5
        'Saturday
        Case 6
            iCount = 6
        'Sunday
        Case 7
            iCount = 7
        Case Else
    End Select
        
    strDayButton = "Day" & iCount & "Button"
    
    For Each ctrlButton In DatePickerForm.Controls
        If Left(ctrlButton.Name, 3) = "Day" And Right(ctrlButton.Name, 6) = "Button" Then
            ctrlButton.Caption = ""
        End If
    Next ctrlButton

    Do Until iCount > 42
        DatePickerForm.Controls(strDayButton).Caption = iDay
        iDay = iDay + 1
        dDate = DateSerial(iYear, iMonth, iDay)
        If Month(dDate) <> Month(dCurrent) Then
            Exit Do
        End If
        iCount = iCount + 1
        strDayButton = "Day" & iCount & "Button"
    Loop
            
End Sub

Function OrdinalSuffix(ByVal lDay As Long) As String

Dim lNumber As Long
Const cSfx = "stndrdthththththth" ' 2 char suffixes

    lNumber = lDay Mod 100
    If ((Abs(lNumber) >= 10) And (Abs(lNumber) <= 19)) _
            Or ((Abs(lNumber) Mod 10) = 0) Then
        OrdinalSuffix = "th"
    Else
        OrdinalSuffix = Mid(cSfx, _
            ((Abs(lNumber) Mod 10) * 2) - 1, 2)
    End If
    
End Function

Function DateFormatting(ByVal dDate As Date) As String

Dim iDay As Integer
Dim strDayName As String
Dim strMonth As String
Dim strSuffix As String

    Select Case DatePart("w", dDate, vbMonday)
        Case 1
            strDayName = "Monday "
        Case 2
            strDayName = "Tuesday "
        Case 3
            strDayName = "Wednesday "
        Case 4
            strDayName = "Thursday "
        Case 5
            strDayName = "Friday "
        Case 6
            strDayName = "Saturday "
        Case 7
            strDayName = "Sunday "
    End Select
    
    iDay = Day(dDate)
    strSuffix = OrdinalSuffix(iDay)
    strMonth = " " & MonthName(Month(dDate), False)
    
    DateFormatting = strDayName & iDay & strSuffix & strMonth

End Function

Function UndoDate(strDate As String) As Date

Dim iCount As Integer: iCount = Len(strDate)
Dim iDay As Integer
Dim iDigit2 As Integer
Dim iMonth As Integer
Dim iYear As Integer
Dim strMonth As String
Dim strTemp As String
Dim vDigit1 As Variant

    Do While iCount > 0
        If Mid(strDate, iCount, 1) = " " And strMonth = "" Then
            strMonth = Right(strDate, Len(strDate) - iCount)
            iMonth = Month(DateValue("03/" & strMonth & "/2014"))
        ElseIf IsNumeric(Mid(strDate, iCount, 1)) Then
            If iDigit2 = 0 And iDay = 0 Then
                iDigit2 = Mid(strDate, iCount, 1)
                vDigit1 = Mid(strDate, iCount - 1, 1)
                If vDigit1 = " " Then
                    iDay = iDigit2
                Else
                    iDay = vDigit1 & iDigit2
                End If
            End If
        End If
        iCount = iCount - 1
    Loop
    
    If iMonth = 12 And Month(Date) = 1 Then
        If iDay >= Day(Date) Then
            iYear = Year(Date) - 1
        Else
            iYear = Year(Date)
        End If
    ElseIf iMonth >= Month(Date) Then
        iYear = Year(Date)
    Else
        iYear = Year(Date) + 1
    End If
    
    If iDay <> 0 Then
        UndoDate = Format(DateValue(iDay & "/" & iMonth & "/" & iYear), "dd/mm/yyyy")
    End If

End Function
