    If nodeId = 0 And expectedNodeId <> 0 Then nodeId = expectedNodeId ' note: this gets stale quickly, but if we know it then don't forget it
    If Not parent Is Nothing Then
        Set parentNode = parent
        If parentId = 0 Then parentId = parent.nodeId
    End If
    If Not b Is Nothing Then
        Set Browser = b
        Set v = Nothing
        If nodeId > 0 Then
            Set v = Browser.cdp.DOM.resolveNode(nodeId)
        ElseIf backendNodeId > 0 Then
            Set v = Browser.cdp.DOM.resolveNode(backendNodeId:=backendNodeId)
        End If
        If Not v Is Nothing Then
            Set resolveNode = New cdpRuntimeRemoteObject
            resolveNode.init v
        End If
    End If

