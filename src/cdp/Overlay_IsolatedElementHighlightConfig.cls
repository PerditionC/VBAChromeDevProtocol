VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayIsolatedElementHighli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "IsolatedElementHighlightConfig"
' Overlay.IsolatedElementHighlightConfig
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' isolationModeHighlightConfig : IsolationModeHighlightConfig
' A descriptor for the highlight appearance of an element in isolation mode.
Public isolationModeHighlightConfig AS cdpOverlayIsolationModeHighligh

' nodeId : DOM.NodeId
' Identifier of the isolated element to highlight.
Public nodeId AS Long

Public Function init(ByVal obj as Dictionary) As cdpOverlayIsolatedElementHighli
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set isolationModeHighlightConfig = New cdpOverlayIsolationModeHighligh
    isolationModeHighlightConfig.init obj.Item("isolationModeHighlightConfig")
    Let nodeId = CLng(obj.Item("nodeId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("isolationModeHighlightConfig") = isolationModeHighlightConfig.asDictionary()
    dict("nodeId") = nodeId

    Set asDictionary = dict
End Function

