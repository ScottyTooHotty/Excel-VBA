Attribute VB_Name = "Postcode_Distance"
Option Explicit

Function G_DISTANCE(strOrigin As String, strDestination As String, bMeasurement As Boolean) As Double

Dim myRequest As XMLHTTP60
Dim myDomDoc As DOMDocument60
Dim distanceNode As IXMLDOMNode

    G_DISTANCE = 0
    
    On Error GoTo exitRoute
    
    strOrigin = Replace(strOrigin, " ", "%20")
    strDestination = Replace(strDestination, " ", "%20")
    
    Set myRequest = New XMLHTTP60
    
    myRequest.Open "GET", "http://maps.googleapis.com/maps/api/directions/xml?origin=" _
        & strOrigin & "&destination=" & strDestination & "&sensor=false", False
    myRequest.send
    
    Set myDomDoc = New DOMDocument60
    
    myDomDoc.LoadXML myRequest.responseText
    
    Set distanceNode = myDomDoc.SelectSingleNode("//leg/distance/value")
    
    If Not distanceNode Is Nothing Then
        If bMeasurement = True Then
            G_DISTANCE = Round(distanceNode.Text / 1000 * 0.621371192, 2)
        Else
            G_DISTANCE = distanceNode.Text / 1000
        End If
    End If
    
exitRoute:
    Set distanceNode = Nothing
    Set myDomDoc = Nothing
    Set myRequest = Nothing
    
End Function


