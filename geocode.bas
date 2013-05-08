Function GoogleGeocode(address As String) As String
    Dim xDoc As New MSXML2.DOMDocument
    xDoc.async = False
    xDoc.Load ("http://maps.googleapis.com/maps/api/geocode/" + _
        "xml?address=" + address + "&sensor=false")
    If xDoc.parseError.ErrorCode <> 0 Then
        GoogleGeocode = xDoc.parseError.reason
    Else
        xDoc.setProperty "SelectionLanguage", "XPath"
        lat = xDoc.SelectSingleNode("//lat").Text
        lng = xDoc.SelectSingleNode("//lng").Text
        GoogleGeocode = lat & "," & lng
    End If
End Function
