VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMBackendNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' DOM.BackendNode
' Backend node with a friendly name.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' nodeType : integer
' `Node`'s nodeType.
Public nodeType AS integer

' nodeName : string
' `Node`'s nodeName.
Public nodeName AS string

' backendNodeId : BackendNodeId
' 
Public backendNodeId AS integer

Public Function init(ByVal obj as Dictionary) As cdpDOMBackendNode

    Let nodeType = CInt(obj.Item("nodeType"))
    Let nodeName = CStr(obj.Item("nodeName"))
    Let backendNodeId = CInt(obj.Item("backendNodeId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("nodeType") = nodeType
    dict("nodeName") = nodeName
    dict("backendNodeId") = backendNodeId

    Set asDictionary = dict
End Function

