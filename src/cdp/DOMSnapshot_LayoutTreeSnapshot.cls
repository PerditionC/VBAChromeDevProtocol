VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotLayoutTreeSnapsho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "LayoutTreeSnapshot"
' DOMSnapshot.LayoutTreeSnapshot
' Table of details of an element in the DOM tree with a LayoutObject.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' nodeIndex : array of integer
' Index of the corresponding node in the `NodeTreeSnapshot` array returned by `captureSnapshot`.
Public nodeIndex AS Collection

' styles : array of cdpDOMSnapshotArrayOfStrings
' Array of indexes specifying computed style strings, filtered according to the `computedStyles` parameter passed to `captureSnapshot`.
Public styles AS Collection

' bounds : array of cdpDOMSnapshotRectangle
' The absolute position bounding box.
Public bounds AS Collection

' text : array of cdpDOMSnapshotStringIndex
' Contents of the LayoutText, if any.
Public text AS Collection

' stackingContexts : RareBooleanData
' Stacking context information.
Public stackingContexts AS cdpDOMSnapshotRareBooleanData

' paintOrders : array of integer
' Optional
' Global paint order index, which is determined by the stacking order of the nodes. Nodes
    '   that are painted together will have the same index. Only provided if includePaintOrder in
    '   captureSnapshot was true.
Public paintOrders AS Collection

' offsetRects : array of cdpDOMSnapshotRectangle
' Optional
' The offset rect of nodes. Only available when includeDOMRects is set to true
Public offsetRects AS Collection

' scrollRects : array of cdpDOMSnapshotRectangle
' Optional
' The scroll rect of nodes. Only available when includeDOMRects is set to true
Public scrollRects AS Collection

' clientRects : array of cdpDOMSnapshotRectangle
' Optional
' The client rect of nodes. Only available when includeDOMRects is set to true
Public clientRects AS Collection

' blendedBackgroundColors : array of cdpDOMSnapshotStringIndex
' Optional
' The list of background colors that are blended with colors of overlapping elements.
Public blendedBackgroundColors AS Collection

' textColorOpacities : array of number
' Optional
' The list of computed text opacities.
Public textColorOpacities AS Collection

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotLayoutTreeSnapsho
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("nodeIndex")
        nodeIndex.Add v
    Next v
    For each v in obj.Item("styles")
        styles.Add v
    Next v
    For each v in obj.Item("bounds")
        bounds.Add v
    Next v
    For each v in obj.Item("text")
        text.Add v
    Next v
    Set stackingContexts = New cdpDOMSnapshotRareBooleanData
    stackingContexts.init obj.Item("stackingContexts")
    If obj.Exists("paintOrders") Then
        For each v in obj.Item("paintOrders")
            paintOrders.Add v
        Next v
    End If
    If obj.Exists("offsetRects") Then
        For each v in obj.Item("offsetRects")
            offsetRects.Add v
        Next v
    End If
    If obj.Exists("scrollRects") Then
        For each v in obj.Item("scrollRects")
            scrollRects.Add v
        Next v
    End If
    If obj.Exists("clientRects") Then
        For each v in obj.Item("clientRects")
            clientRects.Add v
        Next v
    End If
    If obj.Exists("blendedBackgroundColors") Then
        For each v in obj.Item("blendedBackgroundColors")
            blendedBackgroundColors.Add v
        Next v
    End If
    If obj.Exists("textColorOpacities") Then
        For each v in obj.Item("textColorOpacities")
            textColorOpacities.Add v
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("nodeIndex") = nodeIndex
    Set dict("styles") = styles
    Set dict("bounds") = bounds
    Set dict("text") = text
    Set dict("stackingContexts") = stackingContexts.asDictionary()
    Set dict("paintOrders") = paintOrders
    Set dict("offsetRects") = offsetRects
    Set dict("scrollRects") = scrollRects
    Set dict("clientRects") = clientRects
    Set dict("blendedBackgroundColors") = blendedBackgroundColors
    Set dict("textColorOpacities") = textColorOpacities

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set nodeIndex = New Collection
    Set styles = New Collection
    Set bounds = New Collection
    Set text = New Collection
    Set paintOrders = New Collection
    Set offsetRects = New Collection
    Set scrollRects = New Collection
    Set clientRects = New Collection
    Set blendedBackgroundColors = New Collection
    Set textColorOpacities = New Collection
End Sub

