VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpLayerTreeScrollRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScrollRect"
' LayerTree.ScrollRect
' Rectangle where scrolling happens on the main thread.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' rect : DOM.Rect
' Rectangle itself.
Public rect AS cdpDOMRect

' type : string
' Reason for rectangle to force scrolling on the main thread
Public fieldType AS string

Public Function init(ByVal obj as Dictionary) As cdpLayerTreeScrollRect
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set rect = New cdpDOMRect
    rect.init obj.Item("rect")
    Let fieldType = CStr(obj.Item("type"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("rect") = rect.asDictionary()
    dict("type") = fieldType

    Set asDictionary = dict
End Function

