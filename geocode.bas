Public Sub RecalculateSelection()
    ' Disable Automatic Calculations. Then add this as a macro with Ctrl+Q shortcut
    If TypeName(Selection) = "Range" Then
        Selection.Calculate
    End If
End Sub

Function GoogleReverseGeocode(lat As Double, lng As Double) As String
    ' =GoogleReverseGeocode(latitude, longitude) returns an address
    Dim xDoc As New MSXML2.DOMDocument
    xDoc.async = False
    xDoc.Load ("https://maps.googleapis.com/maps/api/geocode/xml?latlng=" + Str(lat) + "," + Str(lng))
    If xDoc.parseError.ErrorCode <> 0 Then
        GoogleReverseGeocode = xDoc.parseError.reason
    Else
        xDoc.SetProperty "SelectionLanguage", "XPath"
        result = xDoc.SelectSingleNode("/GeocodeResponse/status").Text
        If result <> "OK" Then
            GoogleReverseGeocode = result
        Else
            GoogleReverseGeocode = xDoc.SelectSingleNode("//formatted_address").Text
        End If
    End If
End Function

Function GoogleGeocode(address As String) As String
    Const vbErr = 3     ' Red
    Const vbOK = 23     ' Cyan
    Application.Caller.Font.ColorIndex = xlNone
    Dim xDoc As New MSXML2.DOMDocument
    xDoc.async = False
    xDoc.Load ("http://maps.googleapis.com/maps/api/geocode/xml?address=" + address)
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
            lat = xDoc.SelectSingleNode("//lat").Text
            lng = xDoc.SelectSingleNode("//lng").Text
            Application.Caller.Font.ColorIndex = vbOK
            GoogleGeocode = lat & "," & lng
        End If
    End If
End Function

Sub Geocode()
    address = InputBox("Type an address", "Address")
    Dim xDoc As New MSXML2.DOMDocument
    xDoc.async = False
    xDoc.Load ("http://maps.googleapis.com/maps/api/geocode/" + _
        "xml?address=" + address + "&sensor=false")
    If xDoc.parseError.ErrorCode <> 0 Then
        x = xDoc.parseError.reason
    Else
        nRow = 0
        Set xNodes = xDoc.SelectNodes("//*")
        For Each xNode In xNodes
            If xNode.ChildNodes.Length < 2 And xNode.FirstChild.NodeType = NODE_TEXT Then
                nodeName = xNode.nodeName
                Set xParent = xNode.ParentNode
                While Not xParent Is Nothing
                    nodeName = xParent.nodeName & "/" & nodeName
                    Set xParent = xParent.ParentNode
                Wend
                ActiveSheet.Cells(14 + nRow, 1) = nodeName
                ActiveSheet.Cells(14 + nRow, 2) = xNode.Text
                nRow = nRow + 1
            End If
        Next
    End If
End Sub
