VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpOverlay"
' Overlay [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Configuration data for drawing the source order of an elements children.
' SourceOrderConfig AS object

' Configuration data for the highlighting of Grid elements.
' GridHighlightConfig AS object

' Configuration data for the highlighting of Flex container elements.
' FlexContainerHighlightConfig AS object

' Configuration data for the highlighting of Flex item elements.
' FlexItemHighlightConfig AS object

' Style information for drawing a line.
' LineStyle AS object

' Style information for drawing a box.
' BoxStyle AS object

' 
' ContrastAlgorithm AS string
Public Enum ContrastAlgorithm
    CA_aa
    CA_aaa
    CA_apca
End Enum

' Configuration data for the highlighting of page elements.
' HighlightConfig AS object

' 
' ColorFormat AS string
Public Enum ColorFormat
    CF_rgb
    CF_hsl
    CF_hex
End Enum

' Configurations for Persistent Grid Highlight
' GridNodeHighlightConfig AS object

' 
' FlexNodeHighlightConfig AS object

' 
' ScrollSnapContainerHighlightConfig AS object

' 
' ScrollSnapHighlightConfig AS object

' Configuration for dual screen hinge
' HingeConfig AS object

' 
' ContainerQueryHighlightConfig AS object

' 
' ContainerQueryContainerHighlightConfig AS object

' 
' IsolatedElementHighlightConfig AS object

' 
' IsolationModeHighlightConfig AS object

' 
' InspectMode AS string
Public Enum InspectMode
    IM_searchForNode
    IM_searchForUAShadowDOM
    IM_captureAreaScreenshot
    IM_showDistances
    IM_none
End Enum


Public Function ContrastAlgorithmToString(ByVal ca As ContrastAlgorithm) As String
    Dim retVal As String
    Select Case ca
        Case CA_aa
            retVal = "aa"
        Case CA_aaa
            retVal = "aaa"
        Case CA_apca
            retVal = "apca"
        Case Else
            Debug.Print "Warning, unknown value " & ca
    End Select
    ContrastAlgorithmToString = retVal
End Function

Public Function StringToContrastAlgorithm(ByVal s As String) As ContrastAlgorithm
    Dim retVal As ContrastAlgorithm
    Select Case s
        Case "aa"
            retVal = CA_aa
        Case "aaa"
            retVal = CA_aaa
        Case "apca"
            retVal = CA_apca
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToContrastAlgorithm = retVal
End Function


Public Function ColorFormatToString(ByVal cf As ColorFormat) As String
    Dim retVal As String
    Select Case cf
        Case CF_rgb
            retVal = "rgb"
        Case CF_hsl
            retVal = "hsl"
        Case CF_hex
            retVal = "hex"
        Case Else
            Debug.Print "Warning, unknown value " & cf
    End Select
    ColorFormatToString = retVal
End Function

Public Function StringToColorFormat(ByVal s As String) As ColorFormat
    Dim retVal As ColorFormat
    Select Case s
        Case "rgb"
            retVal = CF_rgb
        Case "hsl"
            retVal = CF_hsl
        Case "hex"
            retVal = CF_hex
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToColorFormat = retVal
End Function


Public Function InspectModeToString(ByVal im As InspectMode) As String
    Dim retVal As String
    Select Case im
        Case IM_searchForNode
            retVal = "searchForNode"
        Case IM_searchForUAShadowDOM
            retVal = "searchForUAShadowDOM"
        Case IM_captureAreaScreenshot
            retVal = "captureAreaScreenshot"
        Case IM_showDistances
            retVal = "showDistances"
        Case IM_none
            retVal = "none"
        Case Else
            Debug.Print "Warning, unknown value " & im
    End Select
    InspectModeToString = retVal
End Function

Public Function StringToInspectMode(ByVal s As String) As InspectMode
    Dim retVal As InspectMode
    Select Case s
        Case "searchForNode"
            retVal = IM_searchForNode
        Case "searchForUAShadowDOM"
            retVal = IM_searchForUAShadowDOM
        Case "captureAreaScreenshot"
            retVal = IM_captureAreaScreenshot
        Case "showDistances"
            retVal = IM_showDistances
        Case "none"
            retVal = IM_none
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToInspectMode = retVal
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

' Disables domain notifications.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Overlay.disable", params
End Sub

' Enables domain notifications.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Overlay.enable", params
End Sub

' For testing.
' Returns: highlight - Highlight data for the node.
Public Function getHighlightObjectForTest( _
    ByVal nodeId AS Long, _
    Optional ByVal includeDistance AS Variant, _
    Optional ByVal includeStyle AS Variant, _
    Optional ByVal colorFormat AS Variant, _
    Optional ByVal showAccessibilityInfo AS Variant _
) AS Dictionary
    ' nodeId: DOM.NodeId Id of the node to get highlight object for.
    ' includeDistance: boolean(optional) Whether to include distance info.
    ' includeStyle: boolean(optional) Whether to include style info.
    ' colorFormat: ColorFormat(optional) The color format to get config with (default: hex).
    ' showAccessibilityInfo: boolean(optional) Whether to show accessibility info (default: true).

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    If Not IsMissing(includeDistance) Then params("includeDistance") = CBool(includeDistance)
    If Not IsMissing(includeStyle) Then params("includeStyle") = CBool(includeStyle)
    If Not IsMissing(colorFormat) Then params("colorFormat") = CStr(colorFormat)
    If Not IsMissing(showAccessibilityInfo) Then params("showAccessibilityInfo") = CBool(showAccessibilityInfo)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Overlay.getHighlightObjectForTest", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("highlight") Then Set getHighlightObjectForTest = results("highlight")
    End If
End Function

' For Persistent Grid testing.
' Returns: highlights - Grid Highlight data for the node ids provided.
Public Function getGridHighlightObjectsForTest( _
    ByVal nodeIds AS Collection _
) AS Dictionary
    ' nodeIds: array Ids of the node to get highlight object for.

    Dim params As New Dictionary
    Set params("nodeIds") = nodeIds

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Overlay.getGridHighlightObjectsForTest", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("highlights") Then Set getGridHighlightObjectsForTest = results("highlights")
    End If
End Function

' For Source Order Viewer testing.
' Returns: highlight - Source order highlight data for the node id provided.
Public Function getSourceOrderHighlightObjectForTest( _
    ByVal nodeId AS Long _
) AS Dictionary
    ' nodeId: DOM.NodeId Id of the node to highlight.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Overlay.getSourceOrderHighlightObjectForTest", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("highlight") Then Set getSourceOrderHighlightObjectForTest = results("highlight")
    End If
End Function

' Hides any highlight.
Public Sub hideHighlight()
    Dim params As New Dictionary
    cdp.InvokeMethod "Overlay.hideHighlight", params
End Sub

' Highlights owner element of the frame with given id.
' Deprecated: Doesn't work reliablity and cannot be fixed due to process
' separatation (the owner node might be in a different process). Determine
' the owner node in the client and use highlightNode.
Public Sub highlightFrame( _
    ByVal frameId AS string, _
    Optional ByVal contentColor AS Variant, _
    Optional ByVal contentOutlineColor AS Variant _
)
    ' frameId: Page.FrameId Identifier of the frame to highlight.
    ' contentColor: DOM.RGBA(optional) The content box highlight fill color (default: transparent).
    ' contentOutlineColor: DOM.RGBA(optional) The content box highlight outline color (default: transparent).

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)
    If Not IsMissing(contentColor) Then Set params("contentColor") = contentColor
    If Not IsMissing(contentOutlineColor) Then Set params("contentOutlineColor") = contentOutlineColor

    cdp.InvokeMethod "Overlay.highlightFrame", params
End Sub

' Highlights DOM node with given id or with the given JavaScript object wrapper. Either nodeId or
' objectId must be specified.
Public Sub highlightNode( _
    ByVal highlightConfig AS Dictionary, _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant, _
    Optional ByVal selector AS Variant _
)
    ' highlightConfig: HighlightConfig A descriptor for the highlight appearance.
    ' nodeId: DOM.NodeId(optional) Identifier of the node to highlight.
    ' backendNodeId: DOM.BackendNodeId(optional) Identifier of the backend node to highlight.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node to be highlighted.
    ' selector: string(optional) Selectors to highlight relevant nodes.

    Dim params As New Dictionary
    Set params("highlightConfig") = highlightConfig
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)
    If Not IsMissing(selector) Then params("selector") = CStr(selector)

    cdp.InvokeMethod "Overlay.highlightNode", params
End Sub

' Highlights given quad. Coordinates are absolute with respect to the main frame viewport.
Public Sub highlightQuad( _
    ByVal quad AS Collection, _
    Optional ByVal color AS Variant, _
    Optional ByVal outlineColor AS Variant _
)
    ' quad: DOM.Quad Quad to highlight
    ' color: DOM.RGBA(optional) The highlight fill color (default: transparent).
    ' outlineColor: DOM.RGBA(optional) The highlight outline color (default: transparent).

    Dim params As New Dictionary
    Set params("quad") = quad
    If Not IsMissing(color) Then Set params("color") = color
    If Not IsMissing(outlineColor) Then Set params("outlineColor") = outlineColor

    cdp.InvokeMethod "Overlay.highlightQuad", params
End Sub

' Highlights given rectangle. Coordinates are absolute with respect to the main frame viewport.
Public Sub highlightRect( _
    ByVal x AS Long, _
    ByVal y AS Long, _
    ByVal width AS Long, _
    ByVal height AS Long, _
    Optional ByVal color AS Variant, _
    Optional ByVal outlineColor AS Variant _
)
    ' x: integer X coordinate
    ' y: integer Y coordinate
    ' width: integer Rectangle width
    ' height: integer Rectangle height
    ' color: DOM.RGBA(optional) The highlight fill color (default: transparent).
    ' outlineColor: DOM.RGBA(optional) The highlight outline color (default: transparent).

    Dim params As New Dictionary
    params("x") = CLng(x)
    params("y") = CLng(y)
    params("width") = CLng(width)
    params("height") = CLng(height)
    If Not IsMissing(color) Then Set params("color") = color
    If Not IsMissing(outlineColor) Then Set params("outlineColor") = outlineColor

    cdp.InvokeMethod "Overlay.highlightRect", params
End Sub

' Highlights the source order of the children of the DOM node with given id or with the given
' JavaScript object wrapper. Either nodeId or objectId must be specified.
Public Sub highlightSourceOrder( _
    ByVal sourceOrderConfig AS Dictionary, _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant _
)
    ' sourceOrderConfig: SourceOrderConfig A descriptor for the appearance of the overlay drawing.
    ' nodeId: DOM.NodeId(optional) Identifier of the node to highlight.
    ' backendNodeId: DOM.BackendNodeId(optional) Identifier of the backend node to highlight.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node to be highlighted.

    Dim params As New Dictionary
    Set params("sourceOrderConfig") = sourceOrderConfig
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)

    cdp.InvokeMethod "Overlay.highlightSourceOrder", params
End Sub

' Enters the 'inspect' mode. In this mode, elements that user is hovering over are highlighted.
' Backend then generates 'inspectNodeRequested' event upon element selection.
Public Sub setInspectMode( _
    ByVal mode AS string, _
    Optional ByVal highlightConfig AS Variant _
)
    ' mode: InspectMode Set an inspection mode.
    ' highlightConfig: HighlightConfig(optional) A descriptor for the highlight appearance of hovered-over nodes. May be omitted if `enabled
    '   == false`.

    Dim params As New Dictionary
    params("mode") = CStr(mode)
    If Not IsMissing(highlightConfig) Then Set params("highlightConfig") = highlightConfig

    cdp.InvokeMethod "Overlay.setInspectMode", params
End Sub

' Highlights owner element of all frames detected to be ads.
Public Sub setShowAdHighlights( _
    ByVal show AS boolean _
)
    ' show: boolean True for showing ad highlights

    Dim params As New Dictionary
    params("show") = CBool(show)

    cdp.InvokeMethod "Overlay.setShowAdHighlights", params
End Sub

' 
Public Sub setPausedInDebuggerMessage( _
    Optional ByVal message AS Variant _
)
    ' message: string(optional) The message to display, also triggers resume and step over controls.

    Dim params As New Dictionary
    If Not IsMissing(message) Then params("message") = CStr(message)

    cdp.InvokeMethod "Overlay.setPausedInDebuggerMessage", params
End Sub

' Requests that backend shows debug borders on layers
Public Sub setShowDebugBorders( _
    ByVal show AS boolean _
)
    ' show: boolean True for showing debug borders

    Dim params As New Dictionary
    params("show") = CBool(show)

    cdp.InvokeMethod "Overlay.setShowDebugBorders", params
End Sub

' Requests that backend shows the FPS counter
Public Sub setShowFPSCounter( _
    ByVal show AS boolean _
)
    ' show: boolean True for showing the FPS counter

    Dim params As New Dictionary
    params("show") = CBool(show)

    cdp.InvokeMethod "Overlay.setShowFPSCounter", params
End Sub

' Highlight multiple elements with the CSS Grid overlay.
Public Sub setShowGridOverlays( _
    ByVal gridNodeHighlightConfigs AS Collection _
)
    ' gridNodeHighlightConfigs: array An array of node identifiers and descriptors for the highlight appearance.

    Dim params As New Dictionary
    Set params("gridNodeHighlightConfigs") = gridNodeHighlightConfigs

    cdp.InvokeMethod "Overlay.setShowGridOverlays", params
End Sub

' 
Public Sub setShowFlexOverlays( _
    ByVal flexNodeHighlightConfigs AS Collection _
)
    ' flexNodeHighlightConfigs: array An array of node identifiers and descriptors for the highlight appearance.

    Dim params As New Dictionary
    Set params("flexNodeHighlightConfigs") = flexNodeHighlightConfigs

    cdp.InvokeMethod "Overlay.setShowFlexOverlays", params
End Sub

' 
Public Sub setShowScrollSnapOverlays( _
    ByVal scrollSnapHighlightConfigs AS Collection _
)
    ' scrollSnapHighlightConfigs: array An array of node identifiers and descriptors for the highlight appearance.

    Dim params As New Dictionary
    Set params("scrollSnapHighlightConfigs") = scrollSnapHighlightConfigs

    cdp.InvokeMethod "Overlay.setShowScrollSnapOverlays", params
End Sub

' 
Public Sub setShowContainerQueryOverlays( _
    ByVal containerQueryHighlightConfigs AS Collection _
)
    ' containerQueryHighlightConfigs: array An array of node identifiers and descriptors for the highlight appearance.

    Dim params As New Dictionary
    Set params("containerQueryHighlightConfigs") = containerQueryHighlightConfigs

    cdp.InvokeMethod "Overlay.setShowContainerQueryOverlays", params
End Sub

' Requests that backend shows paint rectangles
Public Sub setShowPaintRects( _
    ByVal result AS boolean _
)
    ' result: boolean True for showing paint rectangles

    Dim params As New Dictionary
    params("result") = CBool(result)

    cdp.InvokeMethod "Overlay.setShowPaintRects", params
End Sub

' Requests that backend shows layout shift regions
Public Sub setShowLayoutShiftRegions( _
    ByVal result AS boolean _
)
    ' result: boolean True for showing layout shift regions

    Dim params As New Dictionary
    params("result") = CBool(result)

    cdp.InvokeMethod "Overlay.setShowLayoutShiftRegions", params
End Sub

' Requests that backend shows scroll bottleneck rects
Public Sub setShowScrollBottleneckRects( _
    ByVal show AS boolean _
)
    ' show: boolean True for showing scroll bottleneck rects

    Dim params As New Dictionary
    params("show") = CBool(show)

    cdp.InvokeMethod "Overlay.setShowScrollBottleneckRects", params
End Sub

' Requests that backend shows hit-test borders on layers
Public Sub setShowHitTestBorders( _
    ByVal show AS boolean _
)
    ' show: boolean True for showing hit-test borders

    Dim params As New Dictionary
    params("show") = CBool(show)

    cdp.InvokeMethod "Overlay.setShowHitTestBorders", params
End Sub

' Request that backend shows an overlay with web vital metrics.
Public Sub setShowWebVitals( _
    ByVal show AS boolean _
)
    ' show: boolean 

    Dim params As New Dictionary
    params("show") = CBool(show)

    cdp.InvokeMethod "Overlay.setShowWebVitals", params
End Sub

' Paints viewport size upon main frame resize.
Public Sub setShowViewportSizeOnResize( _
    ByVal show AS boolean _
)
    ' show: boolean Whether to paint size or not.

    Dim params As New Dictionary
    params("show") = CBool(show)

    cdp.InvokeMethod "Overlay.setShowViewportSizeOnResize", params
End Sub

' Add a dual screen device hinge
Public Sub setShowHinge( _
    Optional ByVal hingeConfig AS Variant _
)
    ' hingeConfig: HingeConfig(optional) hinge data, null means hideHinge

    Dim params As New Dictionary
    If Not IsMissing(hingeConfig) Then Set params("hingeConfig") = hingeConfig

    cdp.InvokeMethod "Overlay.setShowHinge", params
End Sub

' Show elements in isolation mode with overlays.
Public Sub setShowIsolatedElements( _
    ByVal isolatedElementHighlightConfigs AS Collection _
)
    ' isolatedElementHighlightConfigs: array An array of node identifiers and descriptors for the highlight appearance.

    Dim params As New Dictionary
    Set params("isolatedElementHighlightConfigs") = isolatedElementHighlightConfigs

    cdp.InvokeMethod "Overlay.setShowIsolatedElements", params
End Sub

