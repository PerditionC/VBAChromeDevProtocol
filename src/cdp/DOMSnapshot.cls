VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpDOMSnapshot"
' DOMSnapshot [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' A Node in the DOM tree.
' DOMNode AS object

' Details of post layout rendered text positions. The exact layout should not be regarded as
'   stable and may change between versions.
' InlineTextBox AS object

' Details of an element in the DOM tree with a LayoutObject.
' LayoutTreeNode AS object

' A subset of the full ComputedStyle as defined by the request whitelist.
' ComputedStyle AS object

' A name/value pair.
' NameValue AS object

' Index of the string in the strings table.
' StringIndex AS integer

' Index of the string in the strings table.
' ArrayOfStrings AS array

' Data that is only present on rare nodes.
' RareStringData AS object

' 
' RareBooleanData AS object

' 
' RareIntegerData AS object

' 
' Rectangle AS array

' Document snapshot.
' DocumentSnapshot AS object

' Table containing nodes.
' NodeTreeSnapshot AS object

' Table of details of an element in the DOM tree with a LayoutObject.
' LayoutTreeSnapshot AS object

' Table of details of the post layout rendered text positions. The exact layout should not be regarded as
'   stable and may change between versions.
' TextBoxSnapshot AS object



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

' Disables DOM snapshot agent for the given page.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOMSnapshot.disable", params
End Sub

' Enables DOM snapshot agent for the given page.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOMSnapshot.enable", params
End Sub

' Returns a document snapshot, including the full DOM tree of the root node (including iframes,
' template contents, and imported documents) in a flattened array, as well as layout and
' white-listed computed style information for the nodes. Shadow DOM in the returned DOM tree is
' flattened.
' Returns:  - 
Public Function getSnapshot( _
    ByVal computedStyleWhitelist AS Collection, _
    Optional ByVal includeEventListeners AS Variant, _
    Optional ByVal includePaintOrder AS Variant, _
    Optional ByVal includeUserAgentShadowTree AS Variant _
) AS Dictionary
    ' computedStyleWhitelist: array Whitelist of computed styles to return.
    ' includeEventListeners: boolean(optional) Whether or not to retrieve details of DOM listeners (default false).
    ' includePaintOrder: boolean(optional) Whether to determine and include the paint order index of LayoutTreeNodes (default false).
    ' includeUserAgentShadowTree: boolean(optional) Whether to include UA shadow tree in the snapshot (default false).

    Dim params As New Dictionary
    Set params("computedStyleWhitelist") = computedStyleWhitelist
    If Not IsMissing(includeEventListeners) Then params("includeEventListeners") = CBool(includeEventListeners)
    If Not IsMissing(includePaintOrder) Then params("includePaintOrder") = CBool(includePaintOrder)
    If Not IsMissing(includeUserAgentShadowTree) Then params("includeUserAgentShadowTree") = CBool(includeUserAgentShadowTree)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOMSnapshot.getSnapshot", params)
    If cdp.ErrorCode = 0 Then
        Set getSnapshot = results
    End If
End Function

' Returns a document snapshot, including the full DOM tree of the root node (including iframes,
' template contents, and imported documents) in a flattened array, as well as layout and
' white-listed computed style information for the nodes. Shadow DOM in the returned DOM tree is
' flattened.
' Returns:  - 
Public Function captureSnapshot( _
    ByVal computedStyles AS Collection, _
    Optional ByVal includePaintOrder AS Variant, _
    Optional ByVal includeDOMRects AS Variant, _
    Optional ByVal includeBlendedBackgroundColors AS Variant, _
    Optional ByVal includeTextColorOpacities AS Variant _
) AS Dictionary
    ' computedStyles: array Whitelist of computed styles to return.
    ' includePaintOrder: boolean(optional) Whether to include layout object paint orders into the snapshot.
    ' includeDOMRects: boolean(optional) Whether to include DOM rectangles (offsetRects, clientRects, scrollRects) into the snapshot
    ' includeBlendedBackgroundColors: boolean(optional) Whether to include blended background colors in the snapshot (default: false).
    '   Blended background color is achieved by blending background colors of all elements
    '   that overlap with the current element.
    ' includeTextColorOpacities: boolean(optional) Whether to include text color opacity in the snapshot (default: false).
    '   An element might have the opacity property set that affects the text color of the element.
    '   The final text color opacity is computed based on the opacity of all overlapping elements.

    Dim params As New Dictionary
    Set params("computedStyles") = computedStyles
    If Not IsMissing(includePaintOrder) Then params("includePaintOrder") = CBool(includePaintOrder)
    If Not IsMissing(includeDOMRects) Then params("includeDOMRects") = CBool(includeDOMRects)
    If Not IsMissing(includeBlendedBackgroundColors) Then params("includeBlendedBackgroundColors") = CBool(includeBlendedBackgroundColors)
    If Not IsMissing(includeTextColorOpacities) Then params("includeTextColorOpacities") = CBool(includeTextColorOpacities)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOMSnapshot.captureSnapshot", params)
    If cdp.ErrorCode = 0 Then
        Set captureSnapshot = results
    End If
End Function

