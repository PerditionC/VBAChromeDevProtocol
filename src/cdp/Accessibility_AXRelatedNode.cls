VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAccessibilityAXRelatedNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AXRelatedNode"
' Accessibility.AXRelatedNode
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' backendDOMNodeId : DOM.BackendNodeId
' The BackendNodeId of the related DOM node.
Public backendDOMNodeId AS Long

' idref : string
' Optional
' The IDRef value provided, if any.
Public idref AS string

' text : string
' Optional
' The text alternative of this node in the current context.
Public text AS string

Public Function init(ByVal obj as Dictionary) As cdpAccessibilityAXRelatedNode
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let backendDOMNodeId = CLng(obj.Item("backendDOMNodeId"))
    If obj.Exists("idref") Then Let idref = CStr(obj.Item("idref"))
    If obj.Exists("text") Then Let text = CStr(obj.Item("text"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("backendDOMNodeId") = backendDOMNodeId
    if idref <> vbNullString Then dict("idref") = idref
    if text <> vbNullString Then dict("text") = text

    Set asDictionary = dict
End Function

