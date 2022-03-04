VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpDOM"
' DOM
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Unique DOM node identifier.
' NodeId AS integer

' Unique DOM node identifier used to reference a node that may not have been pushed to the
'   front-end.
' BackendNodeId AS integer

' Backend node with a friendly name.
' BackendNode AS object

' Pseudo element type.
' PseudoType AS string
Public Enum PseudoType
    PT_firstline
    PT_firstletter
    PT_before
    PT_after
    PT_marker
    PT_backdrop
    PT_selection
    PT_targettext
    PT_spellingerror
    PT_grammarerror
    PT_highlight
    PT_firstlineinherited
    PT_scrollbar
    PT_scrollbarthumb
    PT_scrollbarbutton
    PT_scrollbartrack
    PT_scrollbartrackpiece
    PT_scrollbarcorner
    PT_resizer
    PT_inputlistbutton
End Enum

' Shadow root type.
' ShadowRootType AS string
Public Enum ShadowRootType
    SRT_useragent
    SRT_open
    SRT_closed
End Enum

' Document compatibility mode.
' CompatibilityMode AS string
Public Enum CompatibilityMode
    CM_QuirksMode
    CM_LimitedQuirksMode
    CM_NoQuirksMode
End Enum

' DOM interaction is implemented in terms of mirror objects that represent the actual DOM nodes.
'   DOMNode is a base node mirror type.
' Node AS object

' A structure holding an RGBA color.
' RGBA AS object

' An array of quad vertices, x immediately followed by y for each point, points clock-wise.
' Quad AS array

' Box model.
' BoxModel AS object

' CSS Shape Outside details.
' ShapeOutsideInfo AS object

' Rectangle.
' Rect AS object

' 
' CSSComputedStyleProperty AS object


Public Function PseudoTypeToString(ByVal pt As PseudoType) As String
    Dim retVal As String
    Select Case pt
        Case PT_firstline
            retVal = "first-line"
        Case PT_firstletter
            retVal = "first-letter"
        Case PT_before
            retVal = "before"
        Case PT_after
            retVal = "after"
        Case PT_marker
            retVal = "marker"
        Case PT_backdrop
            retVal = "backdrop"
        Case PT_selection
            retVal = "selection"
        Case PT_targettext
            retVal = "target-text"
        Case PT_spellingerror
            retVal = "spelling-error"
        Case PT_grammarerror
            retVal = "grammar-error"
        Case PT_highlight
            retVal = "highlight"
        Case PT_firstlineinherited
            retVal = "first-line-inherited"
        Case PT_scrollbar
            retVal = "scrollbar"
        Case PT_scrollbarthumb
            retVal = "scrollbar-thumb"
        Case PT_scrollbarbutton
            retVal = "scrollbar-button"
        Case PT_scrollbartrack
            retVal = "scrollbar-track"
        Case PT_scrollbartrackpiece
            retVal = "scrollbar-track-piece"
        Case PT_scrollbarcorner
            retVal = "scrollbar-corner"
        Case PT_resizer
            retVal = "resizer"
        Case PT_inputlistbutton
            retVal = "input-list-button"
        Case Else
            Debug.Print "Warning, unknown value " & pt
    End Select
    PseudoTypeToString = retVal
End Function

Public Function StringToPseudoType(ByVal s As String) As PseudoType
    Dim retVal As PseudoType
    Select Case s
        Case "first-line"
            retVal = PT_firstline
        Case "first-letter"
            retVal = PT_firstletter
        Case "before"
            retVal = PT_before
        Case "after"
            retVal = PT_after
        Case "marker"
            retVal = PT_marker
        Case "backdrop"
            retVal = PT_backdrop
        Case "selection"
            retVal = PT_selection
        Case "target-text"
            retVal = PT_targettext
        Case "spelling-error"
            retVal = PT_spellingerror
        Case "grammar-error"
            retVal = PT_grammarerror
        Case "highlight"
            retVal = PT_highlight
        Case "first-line-inherited"
            retVal = PT_firstlineinherited
        Case "scrollbar"
            retVal = PT_scrollbar
        Case "scrollbar-thumb"
            retVal = PT_scrollbarthumb
        Case "scrollbar-button"
            retVal = PT_scrollbarbutton
        Case "scrollbar-track"
            retVal = PT_scrollbartrack
        Case "scrollbar-track-piece"
            retVal = PT_scrollbartrackpiece
        Case "scrollbar-corner"
            retVal = PT_scrollbarcorner
        Case "resizer"
            retVal = PT_resizer
        Case "input-list-button"
            retVal = PT_inputlistbutton
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToPseudoType = retVal
End Function


Public Function ShadowRootTypeToString(ByVal srt As ShadowRootType) As String
    Dim retVal As String
    Select Case srt
        Case SRT_useragent
            retVal = "user-agent"
        Case SRT_open
            retVal = "open"
        Case SRT_closed
            retVal = "closed"
        Case Else
            Debug.Print "Warning, unknown value " & srt
    End Select
    ShadowRootTypeToString = retVal
End Function

Public Function StringToShadowRootType(ByVal s As String) As ShadowRootType
    Dim retVal As ShadowRootType
    Select Case s
        Case "user-agent"
            retVal = SRT_useragent
        Case "open"
            retVal = SRT_open
        Case "closed"
            retVal = SRT_closed
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToShadowRootType = retVal
End Function


Public Function CompatibilityModeToString(ByVal cm As CompatibilityMode) As String
    Dim retVal As String
    Select Case cm
        Case CM_QuirksMode
            retVal = "QuirksMode"
        Case CM_LimitedQuirksMode
            retVal = "LimitedQuirksMode"
        Case CM_NoQuirksMode
            retVal = "NoQuirksMode"
        Case Else
            Debug.Print "Warning, unknown value " & cm
    End Select
    CompatibilityModeToString = retVal
End Function

Public Function StringToCompatibilityMode(ByVal s As String) As CompatibilityMode
    Dim retVal As CompatibilityMode
    Select Case s
        Case "QuirksMode"
            retVal = CM_QuirksMode
        Case "LimitedQuirksMode"
            retVal = CM_LimitedQuirksMode
        Case "NoQuirksMode"
            retVal = CM_NoQuirksMode
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCompatibilityMode = retVal
End Function


Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' Collects class names for the node with given id and all of it's child nodes.
' Returns: classNames - Class name list.
' Experimental
Public Function collectClassNamesFromSubtree( _
    ByVal nodeId AS Long _
) AS Collection
    ' nodeId: NodeId Id of the node to collect class names.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.collectClassNamesFromSubtree", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("classNames") Then Set collectClassNamesFromSubtree = results("classNames")
    End If
End Function

' Creates a deep copy of the specified node and places it into the target container before the
' given anchor.
' Returns: nodeId - Id of the node clone.
' Experimental
Public Function copyTo( _
    ByVal nodeId AS Long, _
    ByVal targetNodeId AS Long, _
    Optional ByVal insertBeforeNodeId AS Variant _
) AS Long
    ' nodeId: NodeId Id of the node to copy.
    ' targetNodeId: NodeId Id of the element to drop the copy into.
    ' insertBeforeNodeId: NodeId(optional) Drop the copy before this node (if absent, the copy becomes the last child of
    '   `targetNodeId`).

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("targetNodeId") = CLng(targetNodeId)
    If Not IsMissing(insertBeforeNodeId) Then params("insertBeforeNodeId") = CLng(insertBeforeNodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.copyTo", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeId") Then copyTo = results("nodeId")
    End If
End Function

' Describes node given its id, does not require domain to be enabled. Does not start tracking any
' objects, can be used for automation.
' Returns: node - Node description.
Public Function describeNode( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant, _
    Optional ByVal depth AS Variant, _
    Optional ByVal pierce AS Variant _
) AS Dictionary
    ' nodeId: NodeId(optional) Identifier of the node.
    ' backendNodeId: BackendNodeId(optional) Identifier of the backend node.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper.
    ' depth: integer(optional) The maximum depth at which children should be retrieved, defaults to 1. Use -1 for the
    '   entire subtree or provide an integer larger than 0.
    ' pierce: boolean(optional) Whether or not iframes and shadow roots should be traversed when returning the subtree
    '   (default is false).

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)
    If Not IsMissing(depth) Then params("depth") = CLng(depth)
    If Not IsMissing(pierce) Then params("pierce") = CBool(pierce)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.describeNode", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("node") Then Set describeNode = results("node")
    End If
End Function

' Scrolls the specified rect of the given node into view if not already visible.
' Note: exactly one between nodeId, backendNodeId and objectId should be passed
' to identify the node.
' Experimental
Public Sub scrollIntoViewIfNeeded( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant, _
    Optional ByVal rect AS Variant _
)
    ' nodeId: NodeId(optional) Identifier of the node.
    ' backendNodeId: BackendNodeId(optional) Identifier of the backend node.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper.
    ' rect: Rect(optional) The rect to be scrolled into view, relative to the node's border box, in CSS pixels.
    '   When omitted, center of the node will be used, similar to Element.scrollIntoView.

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)
    If Not IsMissing(rect) Then Set params("rect") = rect

    cdp.InvokeMethod "DOM.scrollIntoViewIfNeeded", params
End Sub

' Disables DOM agent for the given page.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOM.disable", params
End Sub

' Discards search results from the session with the given id. `getSearchResults` should no longer
' be called for that search.
' Experimental
Public Sub discardSearchResults( _
    ByVal searchId AS string _
)
    ' searchId: string Unique search session identifier.

    Dim params As New Dictionary
    params("searchId") = CStr(searchId)

    cdp.InvokeMethod "DOM.discardSearchResults", params
End Sub

' Enables DOM agent for the given page.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOM.enable", params
End Sub

' Focuses the given element.
Public Sub focus( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant _
)
    ' nodeId: NodeId(optional) Identifier of the node.
    ' backendNodeId: BackendNodeId(optional) Identifier of the backend node.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper.

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)

    cdp.InvokeMethod "DOM.focus", params
End Sub

' Returns attributes for the specified node.
' Returns: attributes - An interleaved array of node attribute names and values.
Public Function getAttributes( _
    ByVal nodeId AS Long _
) AS Collection
    ' nodeId: NodeId Id of the node to retrieve attibutes for.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getAttributes", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("attributes") Then Set getAttributes = results("attributes")
    End If
End Function

' Returns boxes for the given node.
' Returns: model - Box model for the node.
Public Function getBoxModel( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant _
) AS Dictionary
    ' nodeId: NodeId(optional) Identifier of the node.
    ' backendNodeId: BackendNodeId(optional) Identifier of the backend node.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper.

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getBoxModel", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("model") Then Set getBoxModel = results("model")
    End If
End Function

' Returns quads that describe node position on the page. This method
' might return multiple quads for inline nodes.
' Returns: quads - Quads that describe node layout relative to viewport.
' Experimental
Public Function getContentQuads( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant _
) AS Collection
    ' nodeId: NodeId(optional) Identifier of the node.
    ' backendNodeId: BackendNodeId(optional) Identifier of the backend node.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper.

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getContentQuads", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("quads") Then Set getContentQuads = results("quads")
    End If
End Function

' Returns the root DOM node (and optionally the subtree) to the caller.
' Returns: root - Resulting node.
Public Function getDocument( _
    Optional ByVal depth AS Variant, _
    Optional ByVal pierce AS Variant _
) AS Dictionary
    ' depth: integer(optional) The maximum depth at which children should be retrieved, defaults to 1. Use -1 for the
    '   entire subtree or provide an integer larger than 0.
    ' pierce: boolean(optional) Whether or not iframes and shadow roots should be traversed when returning the subtree
    '   (default is false).

    Dim params As New Dictionary
    If Not IsMissing(depth) Then params("depth") = CLng(depth)
    If Not IsMissing(pierce) Then params("pierce") = CBool(pierce)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getDocument", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("root") Then Set getDocument = results("root")
    End If
End Function

' Returns the root DOM node (and optionally the subtree) to the caller.
' Deprecated, as it is not designed to work well with the rest of the DOM agent.
' Use DOMSnapshot.captureSnapshot instead.
' Returns: nodes - Resulting node.
Public Function getFlattenedDocument( _
    Optional ByVal depth AS Variant, _
    Optional ByVal pierce AS Variant _
) AS Collection
    ' depth: integer(optional) The maximum depth at which children should be retrieved, defaults to 1. Use -1 for the
    '   entire subtree or provide an integer larger than 0.
    ' pierce: boolean(optional) Whether or not iframes and shadow roots should be traversed when returning the subtree
    '   (default is false).

    Dim params As New Dictionary
    If Not IsMissing(depth) Then params("depth") = CLng(depth)
    If Not IsMissing(pierce) Then params("pierce") = CBool(pierce)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getFlattenedDocument", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodes") Then Set getFlattenedDocument = results("nodes")
    End If
End Function

' Finds nodes with a given computed style in a subtree.
' Returns: nodeIds - Resulting nodes.
' Experimental
Public Function getNodesForSubtreeByStyle( _
    ByVal nodeId AS Long, _
    ByVal computedStyles AS Collection, _
    Optional ByVal pierce AS Variant _
) AS Collection
    ' nodeId: NodeId Node ID pointing to the root of a subtree.
    ' computedStyles: array The style to filter nodes by (includes nodes if any of properties matches).
    ' pierce: boolean(optional) Whether or not iframes and shadow roots in the same target should be traversed when returning the
    '   results (default is false).

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    Set params("computedStyles") = computedStyles
    If Not IsMissing(pierce) Then params("pierce") = CBool(pierce)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getNodesForSubtreeByStyle", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeIds") Then Set getNodesForSubtreeByStyle = results("nodeIds")
    End If
End Function

' Returns node id at given location. Depending on whether DOM domain is enabled, nodeId is
' either returned or not.
' Returns:  - 
Public Function getNodeForLocation( _
    ByVal x AS Long, _
    ByVal y AS Long, _
    Optional ByVal includeUserAgentShadowDOM AS Variant, _
    Optional ByVal ignorePointerEventsNone AS Variant _
) AS Dictionary
    ' x: integer X coordinate.
    ' y: integer Y coordinate.
    ' includeUserAgentShadowDOM: boolean(optional) False to skip to the nearest non-UA shadow root ancestor (default: false).
    ' ignorePointerEventsNone: boolean(optional) Whether to ignore pointer-events: none on elements and hit test them.

    Dim params As New Dictionary
    params("x") = CLng(x)
    params("y") = CLng(y)
    If Not IsMissing(includeUserAgentShadowDOM) Then params("includeUserAgentShadowDOM") = CBool(includeUserAgentShadowDOM)
    If Not IsMissing(ignorePointerEventsNone) Then params("ignorePointerEventsNone") = CBool(ignorePointerEventsNone)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getNodeForLocation", params)
    If cdp.ErrorCode = 0 Then
        Set getNodeForLocation = results
    End If
End Function

' Returns node's HTML markup.
' Returns: outerHTML - Outer HTML markup.
Public Function getOuterHTML( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant _
) AS string
    ' nodeId: NodeId(optional) Identifier of the node.
    ' backendNodeId: BackendNodeId(optional) Identifier of the backend node.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper.

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getOuterHTML", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("outerHTML") Then getOuterHTML = results("outerHTML")
    End If
End Function

' Returns the id of the nearest ancestor that is a relayout boundary.
' Returns: nodeId - Relayout boundary node id for the given node.
' Experimental
Public Function getRelayoutBoundary( _
    ByVal nodeId AS Long _
) AS Long
    ' nodeId: NodeId Id of the node.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getRelayoutBoundary", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeId") Then getRelayoutBoundary = results("nodeId")
    End If
End Function

' Returns search results from given `fromIndex` to given `toIndex` from the search with the given
' identifier.
' Returns: nodeIds - Ids of the search result nodes.
' Experimental
Public Function getSearchResults( _
    ByVal searchId AS string, _
    ByVal fromIndex AS Long, _
    ByVal toIndex AS Long _
) AS Collection
    ' searchId: string Unique search session identifier.
    ' fromIndex: integer Start index of the search result to be returned.
    ' toIndex: integer End index of the search result to be returned.

    Dim params As New Dictionary
    params("searchId") = CStr(searchId)
    params("fromIndex") = CLng(fromIndex)
    params("toIndex") = CLng(toIndex)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getSearchResults", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeIds") Then Set getSearchResults = results("nodeIds")
    End If
End Function

' Hides any highlight.
Public Sub hideHighlight()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOM.hideHighlight", params
End Sub

' Highlights DOM node.
Public Sub highlightNode()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOM.highlightNode", params
End Sub

' Highlights given rectangle.
Public Sub highlightRect()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOM.highlightRect", params
End Sub

' Marks last undoable state.
' Experimental
Public Sub markUndoableState()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOM.markUndoableState", params
End Sub

' Moves node into the new container, places it before the given anchor.
' Returns: nodeId - New id of the moved node.
Public Function moveTo( _
    ByVal nodeId AS Long, _
    ByVal targetNodeId AS Long, _
    Optional ByVal insertBeforeNodeId AS Variant _
) AS Long
    ' nodeId: NodeId Id of the node to move.
    ' targetNodeId: NodeId Id of the element to drop the moved node into.
    ' insertBeforeNodeId: NodeId(optional) Drop node before this one (if absent, the moved node becomes the last child of
    '   `targetNodeId`).

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("targetNodeId") = CLng(targetNodeId)
    If Not IsMissing(insertBeforeNodeId) Then params("insertBeforeNodeId") = CLng(insertBeforeNodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.moveTo", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeId") Then moveTo = results("nodeId")
    End If
End Function

' Searches for a given string in the DOM tree. Use `getSearchResults` to access search results or
' `cancelSearch` to end this search session.
' Returns:  - 
' Experimental
Public Function performSearch( _
    ByVal query AS string, _
    Optional ByVal includeUserAgentShadowDOM AS Variant _
) AS Dictionary
    ' query: string Plain text or query selector or XPath search query.
    ' includeUserAgentShadowDOM: boolean(optional) True to search in user agent shadow DOM.

    Dim params As New Dictionary
    params("query") = CStr(query)
    If Not IsMissing(includeUserAgentShadowDOM) Then params("includeUserAgentShadowDOM") = CBool(includeUserAgentShadowDOM)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.performSearch", params)
    If cdp.ErrorCode = 0 Then
        Set performSearch = results
    End If
End Function

' Requests that the node is sent to the caller given its path. // FIXME, use XPath
' Returns: nodeId - Id of the node for given path.
' Experimental
Public Function pushNodeByPathToFrontend( _
    ByVal path AS string _
) AS Long
    ' path: string Path to node in the proprietary format.

    Dim params As New Dictionary
    params("path") = CStr(path)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.pushNodeByPathToFrontend", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeId") Then pushNodeByPathToFrontend = results("nodeId")
    End If
End Function

' Requests that a batch of nodes is sent to the caller given their backend node ids.
' Returns: nodeIds - The array of ids of pushed nodes that correspond to the backend ids specified in
 '   backendNodeIds.
' Experimental
Public Function pushNodesByBackendIdsToFrontend( _
    ByVal backendNodeIds AS Collection _
) AS Collection
    ' backendNodeIds: array The array of backend node ids.

    Dim params As New Dictionary
    Set params("backendNodeIds") = backendNodeIds

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.pushNodesByBackendIdsToFrontend", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeIds") Then Set pushNodesByBackendIdsToFrontend = results("nodeIds")
    End If
End Function

' Executes `querySelector` on a given node.
' Returns: nodeId - Query selector result.
Public Function querySelector( _
    ByVal nodeId AS Long, _
    ByVal selector AS string _
) AS Long
    ' nodeId: NodeId Id of the node to query upon.
    ' selector: string Selector string.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("selector") = CStr(selector)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.querySelector", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeId") Then querySelector = results("nodeId")
    End If
End Function

' Executes `querySelectorAll` on a given node.
' Returns: nodeIds - Query selector result.
Public Function querySelectorAll( _
    ByVal nodeId AS Long, _
    ByVal selector AS string _
) AS Collection
    ' nodeId: NodeId Id of the node to query upon.
    ' selector: string Selector string.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("selector") = CStr(selector)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.querySelectorAll", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeIds") Then Set querySelectorAll = results("nodeIds")
    End If
End Function

' Re-does the last undone action.
' Experimental
Public Sub redo()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOM.redo", params
End Sub

' Removes attribute with given name from an element with given id.
Public Sub removeAttribute( _
    ByVal nodeId AS Long, _
    ByVal name AS string _
)
    ' nodeId: NodeId Id of the element to remove attribute from.
    ' name: string Name of the attribute to remove.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("name") = CStr(name)

    cdp.InvokeMethod "DOM.removeAttribute", params
End Sub

' Removes node with given id.
Public Sub removeNode( _
    ByVal nodeId AS Long _
)
    ' nodeId: NodeId Id of the node to remove.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    cdp.InvokeMethod "DOM.removeNode", params
End Sub

' Requests that children of the node with given id are returned to the caller in form of
' `setChildNodes` events where not only immediate children are retrieved, but all children down to
' the specified depth.
Public Sub requestChildNodes( _
    ByVal nodeId AS Long, _
    Optional ByVal depth AS Variant, _
    Optional ByVal pierce AS Variant _
)
    ' nodeId: NodeId Id of the node to get children for.
    ' depth: integer(optional) The maximum depth at which children should be retrieved, defaults to 1. Use -1 for the
    '   entire subtree or provide an integer larger than 0.
    ' pierce: boolean(optional) Whether or not iframes and shadow roots should be traversed when returning the sub-tree
    '   (default is false).

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    If Not IsMissing(depth) Then params("depth") = CLng(depth)
    If Not IsMissing(pierce) Then params("pierce") = CBool(pierce)

    cdp.InvokeMethod "DOM.requestChildNodes", params
End Sub

' Requests that the node is sent to the caller given the JavaScript node object reference. All
' nodes that form the path from the node to the root are also sent to the client as a series of
' `setChildNodes` notifications.
' Returns: nodeId - Node id for given object.
Public Function requestNode( _
    ByVal objectId AS string _
) AS Long
    ' objectId: Runtime.RemoteObjectId JavaScript object id to convert into node.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.requestNode", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeId") Then requestNode = results("nodeId")
    End If
End Function

' Resolves the JavaScript node object for a given NodeId or BackendNodeId.
' Returns: object - JavaScript object wrapper for given node.
Public Function resolveNode( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectGroup AS Variant, _
    Optional ByVal executionContextId AS Variant _
) AS Dictionary
    ' nodeId: NodeId(optional) Id of the node to resolve.
    ' backendNodeId: DOM.BackendNodeId(optional) Backend identifier of the node to resolve.
    ' objectGroup: string(optional) Symbolic group name that can be used to release multiple objects.
    ' executionContextId: Runtime.ExecutionContextId(optional) Execution context in which to resolve the node.

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectGroup) Then params("objectGroup") = CStr(objectGroup)
    If Not IsMissing(executionContextId) Then params("executionContextId") = CLng(executionContextId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.resolveNode", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("object") Then Set resolveNode = results("object")
    End If
End Function

' Sets attribute for an element with given id.
Public Sub setAttributeValue( _
    ByVal nodeId AS Long, _
    ByVal name AS string, _
    ByVal value AS string _
)
    ' nodeId: NodeId Id of the element to set attribute for.
    ' name: string Attribute name.
    ' value: string Attribute value.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("name") = CStr(name)
    params("value") = CStr(value)

    cdp.InvokeMethod "DOM.setAttributeValue", params
End Sub

' Sets attributes on element with given id. This method is useful when user edits some existing
' attribute value and types in several attribute name/value pairs.
Public Sub setAttributesAsText( _
    ByVal nodeId AS Long, _
    ByVal text AS string, _
    Optional ByVal name AS Variant _
)
    ' nodeId: NodeId Id of the element to set attributes for.
    ' text: string Text with a number of attributes. Will parse this text using HTML parser.
    ' name: string(optional) Attribute name to replace with new attributes derived from text in case text parsed
    '   successfully.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("text") = CStr(text)
    If Not IsMissing(name) Then params("name") = CStr(name)

    cdp.InvokeMethod "DOM.setAttributesAsText", params
End Sub

' Sets files for the given file input element.
Public Sub setFileInputFiles( _
    ByVal files AS Collection, _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant _
)
    ' files: array Array of file paths to set.
    ' nodeId: NodeId(optional) Identifier of the node.
    ' backendNodeId: BackendNodeId(optional) Identifier of the backend node.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper.

    Dim params As New Dictionary
    Set params("files") = files
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)

    cdp.InvokeMethod "DOM.setFileInputFiles", params
End Sub

' Sets if stack traces should be captured for Nodes. See `Node.getNodeStackTraces`. Default is disabled.
' Experimental
Public Sub setNodeStackTracesEnabled( _
    ByVal enable AS boolean _
)
    ' enable: boolean Enable or disable.

    Dim params As New Dictionary
    params("enable") = CBool(enable)

    cdp.InvokeMethod "DOM.setNodeStackTracesEnabled", params
End Sub

' Gets stack traces associated with a Node. As of now, only provides stack trace for Node creation.
' Returns: creation - Creation stack trace, if available.
' Experimental
Public Function getNodeStackTraces( _
    ByVal nodeId AS Long _
) AS Dictionary
    ' nodeId: NodeId Id of the node to get stack traces for.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getNodeStackTraces", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("creation") Then Set getNodeStackTraces = results("creation")
    End If
End Function

' Returns file information for the given
' File wrapper.
' Returns: path - 
' Experimental
Public Function getFileInfo( _
    ByVal objectId AS string _
) AS string
    ' objectId: Runtime.RemoteObjectId JavaScript object id of the node wrapper.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getFileInfo", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("path") Then getFileInfo = results("path")
    End If
End Function

' Enables console to refer to the node with given id via $x (see Command Line API for more details
' $x functions).
' Experimental
Public Sub setInspectedNode( _
    ByVal nodeId AS Long _
)
    ' nodeId: NodeId DOM node id to be accessible by means of $x command line API.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    cdp.InvokeMethod "DOM.setInspectedNode", params
End Sub

' Sets node name for a node with given id.
' Returns: nodeId - New node's id.
Public Function setNodeName( _
    ByVal nodeId AS Long, _
    ByVal name AS string _
) AS Long
    ' nodeId: NodeId Id of the node to set name for.
    ' name: string New node's name.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("name") = CStr(name)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.setNodeName", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeId") Then setNodeName = results("nodeId")
    End If
End Function

' Sets node value for a node with given id.
Public Sub setNodeValue( _
    ByVal nodeId AS Long, _
    ByVal value AS string _
)
    ' nodeId: NodeId Id of the node to set value for.
    ' value: string New node's value.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("value") = CStr(value)

    cdp.InvokeMethod "DOM.setNodeValue", params
End Sub

' Sets node HTML markup, returns new node id.
Public Sub setOuterHTML( _
    ByVal nodeId AS Long, _
    ByVal outerHTML AS string _
)
    ' nodeId: NodeId Id of the node to set markup for.
    ' outerHTML: string Outer HTML markup to set.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("outerHTML") = CStr(outerHTML)

    cdp.InvokeMethod "DOM.setOuterHTML", params
End Sub

' Undoes the last performed action.
' Experimental
Public Sub undo()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOM.undo", params
End Sub

' Returns iframe node that owns iframe with the given domain.
' Returns:  - 
' Experimental
Public Function getFrameOwner( _
    ByVal frameId AS string _
) AS Dictionary
    ' frameId: Page.FrameId 

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getFrameOwner", params)
    If cdp.ErrorCode = 0 Then
        Set getFrameOwner = results
    End If
End Function

' Returns the container of the given node based on container query conditions.
' If containerName is given, it will find the nearest container with a matching name;
' otherwise it will find the nearest container regardless of its container name.
' Returns: nodeId - The container node for the given node, or null if not found.
' Experimental
Public Function getContainerForNode( _
    ByVal nodeId AS Long, _
    Optional ByVal containerName AS Variant _
) AS Long
    ' nodeId: NodeId 
    ' containerName: string(optional) 

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    If Not IsMissing(containerName) Then params("containerName") = CStr(containerName)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getContainerForNode", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeId") Then getContainerForNode = results("nodeId")
    End If
End Function

' Returns the descendants of a container query container that have
' container queries against this container.
' Returns: nodeIds - Descendant nodes with container queries against the given container.
' Experimental
Public Function getQueryingDescendantsForContainer( _
    ByVal nodeId AS Long _
) AS Collection
    ' nodeId: NodeId Id of the container node to find querying descendants from.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOM.getQueryingDescendantsForContainer", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeIds") Then Set getQueryingDescendantsForContainer = results("nodeIds")
    End If
End Function

