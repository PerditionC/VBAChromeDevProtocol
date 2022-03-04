VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayGridNodeHighlightConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "GridNodeHighlightConfig"
' Overlay.GridNodeHighlightConfig
' Configurations for Persistent Grid Highlight
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' gridHighlightConfig : GridHighlightConfig
' A descriptor for the highlight appearance.
Public gridHighlightConfig AS cdpOverlayGridHighlightConfig

' nodeId : DOM.NodeId
' Identifier of the node to highlight.
Public nodeId AS Long

Public Function init(ByVal obj as Dictionary) As cdpOverlayGridNodeHighlightConf
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set gridHighlightConfig = New cdpOverlayGridHighlightConfig
    gridHighlightConfig.init obj.Item("gridHighlightConfig")
    Let nodeId = CLng(obj.Item("nodeId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("gridHighlightConfig") = gridHighlightConfig.asDictionary()
    dict("nodeId") = nodeId

    Set asDictionary = dict
End Function

