VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayScrollSnapHighlightCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScrollSnapHighlightConfig"
' Overlay.ScrollSnapHighlightConfig
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' scrollSnapContainerHighlightConfig : ScrollSnapContainerHighlightConfig
' A descriptor for the highlight appearance of scroll snap containers.
Public scrollSnapContainerHighlightConfig AS cdpOverlayScrollSnapContainerHi

' nodeId : DOM.NodeId
' Identifier of the node to highlight.
Public nodeId AS Long

Public Function init(ByVal obj as Dictionary) As cdpOverlayScrollSnapHighlightCo
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set scrollSnapContainerHighlightConfig = New cdpOverlayScrollSnapContainerHi
    scrollSnapContainerHighlightConfig.init obj.Item("scrollSnapContainerHighlightConfig")
    Let nodeId = CLng(obj.Item("nodeId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("scrollSnapContainerHighlightConfig") = scrollSnapContainerHighlightConfig.asDictionary()
    dict("nodeId") = nodeId

    Set asDictionary = dict
End Function

