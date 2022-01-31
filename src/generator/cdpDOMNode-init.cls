    If Not b Is Nothing Then 
        Set browser = b
        If nodeId > 0 Then Set resolveNode = browser.cdp.DOM.resolveNode(nodeId)
        If nodeId > 0 Then Set objectProperties = browser.cdp.Runtime.getProperties(resolveNode("objectId"))
    End If

