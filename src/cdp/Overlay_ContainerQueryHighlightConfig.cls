VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayContainerQueryHighlig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ContainerQueryHighlightConfig"
' Overlay.ContainerQueryHighlightConfig
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' containerQueryContainerHighlightConfig : ContainerQueryContainerHighlightConfig
' A descriptor for the highlight appearance of container query containers.
Public containerQueryContainerHighlightConfig AS cdpOverlayContainerQueryContain

' nodeId : DOM.NodeId
' Identifier of the container node to highlight.
Public nodeId AS Long

Public Function init(ByVal obj as Dictionary) As cdpOverlayContainerQueryHighlig
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set containerQueryContainerHighlightConfig = New cdpOverlayContainerQueryContain
    containerQueryContainerHighlightConfig.init obj.Item("containerQueryContainerHighlightConfig")
    Let nodeId = CLng(obj.Item("nodeId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("containerQueryContainerHighlightConfig") = containerQueryContainerHighlightConfig.asDictionary()
    dict("nodeId") = nodeId

    Set asDictionary = dict
End Function

