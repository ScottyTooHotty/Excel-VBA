Attribute VB_Name = "OutlookUWMT"
Option Explicit

Const xlUp As Long = -4162
'IMPORTANT!!
'Change this string to change which folder and worksheet to use
Const strSheet As String = "20160314"

Sub ExportToExcel(MyMail As MailItem)

Dim lRow As Long
Dim oXLApp As Object
Dim oXLwb As Object
Dim oXLws As Object
Dim olMail As Outlook.MailItem
Dim olNS As Outlook.Namespace
Dim strDataAccessed As String
Dim strFileName As String
Dim strID As String
Dim strMyAr() As String
    
    'Specify which email to look at
    strID = MyMail.EntryID
    Set olNS = Application.GetNamespace("MAPI")
    Set olMail = olNS.GetItemFromID(strID)
    
    'Pick out which data was accessed from the email body
    strMyAr = Split(olMail.Body, vbCrLf)
    strDataAccessed = Right(strMyAr(5), Len(strMyAr(5)) - 15)

    'Establish an Excel application object
    'Make sure Excel and UWMT Data Collection Workbook are already open
    Set oXLApp = GetObject(, "Excel.Application")

    oXLApp.Visible = True
    
    'Set the workbook and worksheet
    Set oXLwb = oXLApp.Workbooks("UWMT Data Collection")
    Set oXLws = oXLwb.Sheets(strSheet)

    'Find the last row
    lRow = oXLws.Range("A" & oXLApp.Rows.Count).End(xlUp).Row + 1

    'Output email data to excel
    With oXLws
        .Range("A" & lRow).Value = olMail.SenderName
        .Range("C" & lRow).Value = strDataAccessed
        .Range("D" & lRow).Value = olMail.ReceivedTime
    End With

    'Reset objects to free up memory
    Set oXLws = Nothing
    Set oXLwb = Nothing
    Set oXLApp = Nothing
    Set olMail = Nothing
    Set olNS = Nothing
    
End Sub
