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
