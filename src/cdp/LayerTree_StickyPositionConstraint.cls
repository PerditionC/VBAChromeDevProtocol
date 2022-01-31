VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpLayerTreeStickyPositionConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "StickyPositionConstraint"
' LayerTree.StickyPositionConstraint
' Sticky position constraints.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' stickyBoxRect : DOM.Rect
' Layout rectangle of the sticky element before being shifted
Public stickyBoxRect AS cdpDOMRect

' containingBlockRect : DOM.Rect
' Layout rectangle of the containing block of the sticky element
Public containingBlockRect AS cdpDOMRect

' nearestLayerShiftingStickyBox : LayerId
' Optional
' The nearest sticky layer that shifts the sticky box
Public nearestLayerShiftingStickyBox AS string

' nearestLayerShiftingContainingBlock : LayerId
' Optional
' The nearest sticky layer that shifts the containing block
Public nearestLayerShiftingContainingBlock AS string

Public Function init(ByVal obj as Dictionary) As cdpLayerTreeStickyPositionConst
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set stickyBoxRect = New cdpDOMRect
    stickyBoxRect.init obj.Item("stickyBoxRect")
    Set containingBlockRect = New cdpDOMRect
    containingBlockRect.init obj.Item("containingBlockRect")
    If obj.Exists("nearestLayerShiftingStickyBox") Then Let nearestLayerShiftingStickyBox = CStr(obj.Item("nearestLayerShiftingStickyBox"))
    If obj.Exists("nearestLayerShiftingContainingBlock") Then Let nearestLayerShiftingContainingBlock = CStr(obj.Item("nearestLayerShiftingContainingBlock"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("stickyBoxRect") = stickyBoxRect.asDictionary()
    Set dict("containingBlockRect") = containingBlockRect.asDictionary()
    if nearestLayerShiftingStickyBox <> vbNullString Then dict("nearestLayerShiftingStickyBox") = nearestLayerShiftingStickyBox
    if nearestLayerShiftingContainingBlock <> vbNullString Then dict("nearestLayerShiftingContainingBlock") = nearestLayerShiftingContainingBlock

    Set asDictionary = dict
End Function

