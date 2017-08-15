VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePickerForm 
   Caption         =   "Date Picker"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "DatePickerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    
    Unload Me
    
End Sub

Private Sub ConfirmButton_Click()

    If DateTextBox.Text = "" Then
        Exit Sub
    ElseIf IsDate(DateTextBox.Text) Then
        dFinalDate = Format(DateTextBox.Text, "dd/mm/yyyy")
        ActiveCell.Value = DateFormatting(DateTextBox.Text)
        Unload Me
    Else
        Exit Sub
    End If

End Sub

Private Sub Day1Button_Click()

Dim dDate As Date

    If Day1Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day1Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day2Button_Click()

Dim dDate As Date

    If Day2Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day2Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day3Button_Click()

Dim dDate As Date

    If Day3Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day3Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day4Button_Click()

Dim dDate As Date

    If Day4Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day4Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day5Button_Click()

Dim dDate As Date

    If Day5Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day5Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day6Button_Click()

Dim dDate As Date

    If Day6Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day6Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day7Button_Click()

Dim dDate As Date

    If Day7Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day7Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day8Button_Click()

Dim dDate As Date

    If Day8Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day8Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day9Button_Click()

Dim dDate As Date

    If Day9Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day9Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day10Button_Click()

Dim dDate As Date

    If Day10Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day10Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day11Button_Click()

Dim dDate As Date

    If Day11Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day11Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day12Button_Click()

Dim dDate As Date

    If Day12Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day12Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day13Button_Click()

Dim dDate As Date

    If Day13Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day13Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day14Button_Click()

Dim dDate As Date

    If Day14Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day14Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day15Button_Click()

Dim dDate As Date

    If Day15Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day15Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day16Button_Click()

Dim dDate As Date

    If Day16Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day16Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day17Button_Click()

Dim dDate As Date

    If Day17Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day17Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day18Button_Click()

Dim dDate As Date

    If Day18Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day18Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day19Button_Click()

Dim dDate As Date

    If Day19Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day19Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day20Button_Click()

Dim dDate As Date

    If Day20Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day20Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day21Button_Click()

Dim dDate As Date

    If Day21Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day21Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day22Button_Click()

Dim dDate As Date

    If Day22Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day22Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day23Button_Click()

Dim dDate As Date

    If Day23Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day23Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day24Button_Click()

Dim dDate As Date

    If Day24Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day24Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day25Button_Click()

Dim dDate As Date

    If Day25Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day25Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day26Button_Click()

Dim dDate As Date

    If Day26Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day26Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day27Button_Click()

Dim dDate As Date

    If Day27Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day27Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day28Button_Click()

Dim dDate As Date

    If Day28Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day28Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day29Button_Click()

Dim dDate As Date

    If Day29Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day29Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day30Button_Click()

Dim dDate As Date

    If Day30Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day30Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day31Button_Click()

Dim dDate As Date

    If Day31Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day31Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day32Button_Click()

Dim dDate As Date

    If Day32Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day32Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day33Button_Click()

Dim dDate As Date

    If Day33Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day33Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day34Button_Click()

Dim dDate As Date

    If Day34Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day34Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day35Button_Click()

Dim dDate As Date

    If Day35Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day35Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day36Button_Click()

Dim dDate As Date

    If Day36Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day36Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day37Button_Click()

Dim dDate As Date

    If Day37Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day37Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day38Button_Click()

Dim dDate As Date

    If Day38Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day38Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day39Button_Click()

Dim dDate As Date

    If Day39Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day39Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day40Button_Click()

Dim dDate As Date

    If Day40Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day40Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day41Button_Click()

Dim dDate As Date

    If Day41Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day41Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub Day42Button_Click()

Dim dDate As Date

    If Day42Button.Caption = "" Then
        Exit Sub
    Else
        dDate = Format(DateAdd("d", CInt(Day42Button.Caption), _
            Format(MonthYearLabel.Caption, "dd/mm/yyyy")) - 1, "dd/mm/yyyy")
        DateTextBox.Text = Format(dDate, "dd/mm/yyyy")
    End If

End Sub

Private Sub NextMonthButton_Click()

    MonthYearLabel.Caption = Format(DateAdd("m", 1, MonthYearLabel.Caption), "mmm-yyyy")
    Call GenerateCalendar(MonthYearLabel.Caption)

End Sub

Private Sub NextYearButton_Click()

    MonthYearLabel.Caption = Format(DateAdd("yyyy", 1, MonthYearLabel.Caption), "mmm-yyyy")
    Call GenerateCalendar(MonthYearLabel.Caption)

End Sub

Private Sub PreviousMonthButton_Click()

    MonthYearLabel.Caption = Format(DateAdd("m", -1, MonthYearLabel.Caption), "mmm-yyyy")
    Call GenerateCalendar(MonthYearLabel.Caption)

End Sub

Private Sub PreviousYearButton_Click()

    MonthYearLabel.Caption = Format(DateAdd("yyyy", -1, MonthYearLabel.Caption), "mmm-yyyy")
    Call GenerateCalendar(MonthYearLabel.Caption)

End Sub

Private Sub UserForm_Activate()
     
Dim TopOffset As Integer
Dim LeftOffset As Integer
     
    TopOffset = (Application.UsableHeight / 2) - (Me.Height / 2)
    LeftOffset = (Application.UsableWidth / 2) - (Me.Width / 2)
     
    Me.Top = Application.Top + TopOffset
    Me.Left = Application.Left + LeftOffset

    If IsDate(ActiveCell.Value) Then
        MonthYearLabel.Caption = Format(ActiveCell.Value, "mmm-yyyy")
        DateTextBox.Text = Format(ActiveCell.Value, "dd/mm/yyyy")
    Else
        MonthYearLabel.Caption = Format(Date, "mmm-yyyy")
        DateTextBox.Text = Format(Date, "dd/mm/yyyy")
    End If
    
    Call GenerateCalendar(MonthYearLabel.Caption)

End Sub

