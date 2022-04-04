Const vbErr = 3     ' Red
Const vbOK = 23     ' Cyan

Public Sub RecalculateSelection()
    If TypeName(Selection) = "Range" Then
        Selection.Calculate
    End If
End Sub

Function GoogleReverseGeocode(lat As Double, lng As Double, key As String) As String
    Dim xDoc As New MSXML2.DOMDocument
    xDoc.async = False
    xDoc.Load ("https://maps.googleapis.com/maps/api/geocode/xml?key=" & key & "&latlng=" & lat + "," & lng)
    If xDoc.parseError.ErrorCode <> 0 Then
        Application.Caller.Font.ColorIndex = vbErr
        GoogleReverseGeocode = xDoc.parseError.reason
    Else
        xDoc.SetProperty "SelectionLanguage", "XPath"
        result = xDoc.SelectSingleNode("/GeocodeResponse/status").Text
        If result <> "OK" Then
            Application.Caller.Font.ColorIndex = vbErr
            GoogleReverseGeocode = result
        Else
            Application.Caller.Font.ColorIndex = vbOK
            GoogleReverseGeocode = xDoc.SelectSingleNode("//formatted_address").Text
        End If
    End If
End Function

Function GoogleGeocode(address As String, key As String) As String
    Application.Caller.Font.ColorIndex = xlNone
    Dim xDoc As New MSXML2.DOMDocument
    xDoc.async = False
    xDoc.Load ("https://maps.googleapis.com/maps/api/geocode/xml?key=" + key + "&address=" + WorksheetFunction.EncodeURL(address))
    If xDoc.parseError.ErrorCode <> 0 Then
        Application.Caller.Font.ColorIndex = vbErr
        GoogleGeocode = xDoc.parseError.reason
    Else
        xDoc.SetProperty "SelectionLanguage", "XPath"
        result = xDoc.SelectSingleNode("/GeocodeResponse/status").Text
        If result <> "OK" Then
            Application.Caller.Font.ColorIndex = vbErr
            GoogleGeocode = result
        Else
            Application.Caller.Font.ColorIndex = vbOK
            GoogleGeocode = xDoc.SelectSingleNode("//lat").Text & "," & xDoc.SelectSingleNode("//lng").Text
        End If
    End If
End Function

Function NominatimReverseGeocode(lat As Double, lng As Double) As String
    On Error GoTo eh
    Dim xDoc As New MSXML2.DOMDocument
    xDoc.async = False
    Url = "https://nominatim.openstreetmap.org/reverse?lat=" & lat & "&lon=" & lng
    xDoc.Load (Url)
    If xDoc.parseError.ErrorCode <> 0 Then
        Application.Caller.Font.ColorIndex = vbErr
        NominatimReverseGeocode = xDoc.parseError.reason
    Else
        xDoc.SetProperty "SelectionLanguage", "XPath"
        Dim loc As MSXML2.IXMLDOMElement
        Set loc = xDoc.SelectSingleNode("/reversegeocode/result")
        If loc Is Nothing Then
            Application.Caller.Font.ColorIndex = vbErr
            NominatimReverseGeocode = xDoc.XML
        Else
            Application.Caller.Font.ColorIndex = vbOK
            NominatimReverseGeocode = loc.Text
        End If
    End If
    Exit Function
eh:
    Debug.Print err.Description
End Function


Function NominatimGeocode(address As String) As String
    Application.Caller.Font.ColorIndex = xlNone
    Dim xDoc As New MSXML2.DOMDocument
    xDoc.async = False
    xDoc.Load ("https://nominatim.openstreetmap.org/search?format=xml&q=" + WorksheetFunction.EncodeURL(address))
    If xDoc.parseError.ErrorCode <> 0 Then
        Application.Caller.Font.ColorIndex = vbErr
        NominatimGeocode = xDoc.parseError.reason
    Else
        xDoc.SetProperty "SelectionLanguage", "XPath"
        Dim loc As MSXML2.IXMLDOMElement
        Set loc = xDoc.SelectSingleNode("/searchresults/place")
        If loc Is Nothing Then
            Application.Caller.Font.ColorIndex = vbErr
            NominatimGeocode = xDoc.XML
        Else
            Application.Caller.Font.ColorIndex = vbOK
            NominatimGeocode = loc.getAttribute("lat") & "," & loc.getAttribute("lon")
        End If
    End If
End Function
