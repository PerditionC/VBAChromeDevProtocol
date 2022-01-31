VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlaySourceOrderConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SourceOrderConfig"
' Overlay.SourceOrderConfig
' Configuration data for drawing the source order of an elements children.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' parentOutlineColor : DOM.RGBA
' the color to outline the givent element in.
Public parentOutlineColor AS cdpDOMRGBA

' childOutlineColor : DOM.RGBA
' the color to outline the child elements in.
Public childOutlineColor AS cdpDOMRGBA

Public Function init(ByVal obj as Dictionary) As cdpOverlaySourceOrderConfig
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set parentOutlineColor = New cdpDOMRGBA
    parentOutlineColor.init obj.Item("parentOutlineColor")
    Set childOutlineColor = New cdpDOMRGBA
    childOutlineColor.init obj.Item("childOutlineColor")

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("parentOutlineColor") = parentOutlineColor.asDictionary()
    Set dict("childOutlineColor") = childOutlineColor.asDictionary()

    Set asDictionary = dict
End Function

