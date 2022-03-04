VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpLayerTreeLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Layer"
' LayerTree.Layer
' Information about a compositing layer.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' layerId : LayerId
' The unique id for this layer.
Public layerId AS string

' parentLayerId : LayerId
' Optional
' The id of parent (not present for root).
Public parentLayerId AS string

' backendNodeId : DOM.BackendNodeId
' Optional
' The backend id for the node associated with this layer.
Public backendNodeId AS Long

' offsetX : number
' Offset from parent layer, X coordinate.
Public offsetX AS Double

' offsetY : number
' Offset from parent layer, Y coordinate.
Public offsetY AS Double

' width : number
' Layer width.
Public width AS Double

' height : number
' Layer height.
Public height AS Double

' transform : array of number
' Optional
' Transformation matrix for layer, default is identity matrix
Public transform AS Collection

' anchorX : number
' Optional
' Transform anchor point X, absent if no transform specified
Public anchorX AS Double

' anchorY : number
' Optional
' Transform anchor point Y, absent if no transform specified
Public anchorY AS Double

' anchorZ : number
' Optional
' Transform anchor point Z, absent if no transform specified
Public anchorZ AS Double

' paintCount : integer
' Indicates how many time this layer has painted.
Public paintCount AS Long

' drawsContent : boolean
' Indicates whether this layer hosts any content, rather than being used for
    '   transform/scrolling purposes only.
Public drawsContent AS boolean

' invisible : boolean
' Optional
' Set if layer is not visible.
Public invisible AS Variant ' boolean

' scrollRects : array of cdpLayerTreeScrollRect
' Optional
' Rectangles scrolling on main thread only.
Public scrollRects AS Collection

' stickyPositionConstraint : StickyPositionConstraint
' Optional
' Sticky position constraint information
Public stickyPositionConstraint AS cdpLayerTreeStickyPositionConst

Public Function init(ByVal obj as Dictionary) As cdpLayerTreeLayer
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let layerId = CStr(obj.Item("layerId"))
    If obj.Exists("parentLayerId") Then Let parentLayerId = CStr(obj.Item("parentLayerId"))
    If obj.Exists("backendNodeId") Then Let backendNodeId = CLng(obj.Item("backendNodeId"))
    Let offsetX = CDbl(obj.Item("offsetX"))
    Let offsetY = CDbl(obj.Item("offsetY"))
    Let width = CDbl(obj.Item("width"))
    Let height = CDbl(obj.Item("height"))
    If obj.Exists("transform") Then
        For each v in obj.Item("transform")
            transform.Add v
        Next v
    End If
    If obj.Exists("anchorX") Then Let anchorX = CDbl(obj.Item("anchorX"))
    If obj.Exists("anchorY") Then Let anchorY = CDbl(obj.Item("anchorY"))
    If obj.Exists("anchorZ") Then Let anchorZ = CDbl(obj.Item("anchorZ"))
    Let paintCount = CLng(obj.Item("paintCount"))
    Let drawsContent = CBool(obj.Item("drawsContent"))
    If obj.Exists("invisible") Then Let invisible = CBool(obj.Item("invisible"))
    If obj.Exists("scrollRects") Then
        For each v in obj.Item("scrollRects")
            Dim obj_scrollRects As cdpLayerTreeScrollRect: Set obj_scrollRects = New cdpLayerTreeScrollRect
            obj_scrollRects.init v
            scrollRects.Add obj_scrollRects
            Set obj_scrollRects = Nothing
        Next v
    End If
    If obj.Exists("stickyPositionConstraint") Then
        Set stickyPositionConstraint = New cdpLayerTreeStickyPositionConst
        stickyPositionConstraint.init obj.Item("stickyPositionConstraint")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("layerId") = layerId
    if parentLayerId <> vbNullString Then dict("parentLayerId") = parentLayerId
    dict("backendNodeId") = backendNodeId
    dict("offsetX") = offsetX
    dict("offsetY") = offsetY
    dict("width") = width
    dict("height") = height
    Set dict("transform") = transform
    dict("anchorX") = anchorX
    dict("anchorY") = anchorY
    dict("anchorZ") = anchorZ
    dict("paintCount") = paintCount
    dict("drawsContent") = drawsContent
    if Not IsEmpty(invisible) Then dict("invisible") = invisible
    Set dict("scrollRects") = scrollRects
    If Not stickyPositionConstraint Is Nothing Then Set dict("stickyPositionConstraint") = stickyPositionConstraint.asDictionary()

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set transform = New Collection
    Set scrollRects = New Collection
End Sub

