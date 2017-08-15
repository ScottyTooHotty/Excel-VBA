Attribute VB_Name = "CDO_Mailer"
Option Explicit

Sub Auto_CDO_Mail(strErrorNum As String, strErrorDesc As String, strFilePath As String, _
    strProcedure As String, strUsername As String)

Dim oMsg As Object
Dim oConfig As Object
Dim strBody As String
Dim vFields As Variant

    On Error GoTo Terminate

    Set oMsg = CreateObject("CDO.Message")
    Set oConfig = CreateObject("CDO.Configuration")

    oConfig.Load -1    ' CDO Source Defaults
    Set vFields = oConfig.Fields
    With vFields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
            = "cas.ccc.cambridgeshire.gov.uk"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Update
    End With

    strBody = "Please help, I broke the workbook." & vbNewLine & vbNewLine & _
        "Error Number: " & strErrorNum & vbNewLine & "Error Description: " & strErrorDesc & _
        vbNewLine & "ActiveWorksheet: " & ActiveSheet.Name & vbNewLine & "ActiveCell: " & _
        ActiveCell.Address & vbNewLine & "Selection: " & Selection.Address & vbNewLine & _
        "Procedure: " & strProcedure & vbNewLine & "File: <<" & strFilePath & ">>" & _
        vbNewLine & vbNewLine & "Thanks" & vbNewLine & strUsername

    With oMsg
        Set .Configuration = oConfig
        .to = "email@address.co.uk"
        .CC = ""
        .BCC = ""
        .From = strUsername & " <auto_email@email.co.uk>" 'Edit name here
        .Subject = "Workbook Error"
        .TextBody = strBody
        '.AddAttachment ("C:\Users\Public\Error - " & Format(Date, "yyyymmdd") & ".jpg")
        .send
    End With

    'Kill "C:\Users\Public\Error - " & Format(Date, "yyyymmdd") & ".jpg"

Terminate:

End Sub

