VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAccessibilityAXNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AXNode"
' Accessibility.AXNode
' A node in the accessibility tree.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' nodeId : AXNodeId
' Unique identifier for this node.
Public nodeId AS string

' ignored : boolean
' Whether this node is ignored for accessibility
Public ignored AS boolean

' ignoredReasons : array of cdpAccessibilityAXProperty
' Optional
' Collection of reasons why this node is hidden.
Public ignoredReasons AS Collection

' role : AXValue
' Optional
' This `Node`'s role, whether explicit or implicit.
Public role AS cdpAccessibilityAXValue

' name : AXValue
' Optional
' The accessible name for this `Node`.
Public name AS cdpAccessibilityAXValue

' description : AXValue
' Optional
' The accessible description for this `Node`.
Public description AS cdpAccessibilityAXValue

' value : AXValue
' Optional
' The value for this `Node`.
Public value AS cdpAccessibilityAXValue

' properties : array of cdpAccessibilityAXProperty
' Optional
' All other properties
Public properties AS Collection

' childIds : array of cdpAccessibilityAXNodeId
' Optional
' IDs for each of this node's child nodes.
Public childIds AS Collection

' backendDOMNodeId : DOM.BackendNodeId
' Optional
' The backend ID for the associated DOM node, if any.
Public backendDOMNodeId AS Long

Public Function init(ByVal obj as Dictionary) As cdpAccessibilityAXNode
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let nodeId = CStr(obj.Item("nodeId"))
    Let ignored = CBool(obj.Item("ignored"))
    If obj.Exists("ignoredReasons") Then
        For each v in obj.Item("ignoredReasons")
            Dim obj_ignoredReasons As cdpAccessibilityAXProperty: Set obj_ignoredReasons = New cdpAccessibilityAXProperty
            obj_ignoredReasons.init v
            ignoredReasons.Add obj_ignoredReasons
            Set obj_ignoredReasons = Nothing
        Next v
    End If
    If obj.Exists("role") Then
        Set role = New cdpAccessibilityAXValue
        role.init obj.Item("role")
    End If
    If obj.Exists("name") Then
        Set name = New cdpAccessibilityAXValue
        name.init obj.Item("name")
    End If
    If obj.Exists("description") Then
        Set description = New cdpAccessibilityAXValue
        description.init obj.Item("description")
    End If
    If obj.Exists("value") Then
        Set value = New cdpAccessibilityAXValue
        value.init obj.Item("value")
    End If
    If obj.Exists("properties") Then
        For each v in obj.Item("properties")
            Dim obj_properties As cdpAccessibilityAXProperty: Set obj_properties = New cdpAccessibilityAXProperty
            obj_properties.init v
            properties.Add obj_properties
            Set obj_properties = Nothing
        Next v
    End If
    If obj.Exists("childIds") Then
        For each v in obj.Item("childIds")
            childIds.Add v
        Next v
    End If
    If obj.Exists("backendDOMNodeId") Then Let backendDOMNodeId = CLng(obj.Item("backendDOMNodeId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("nodeId") = nodeId
    dict("ignored") = ignored
    Set dict("ignoredReasons") = ignoredReasons
    If Not role Is Nothing Then Set dict("role") = role.asDictionary()
    If Not name Is Nothing Then Set dict("name") = name.asDictionary()
    If Not description Is Nothing Then Set dict("description") = description.asDictionary()
    If Not value Is Nothing Then Set dict("value") = value.asDictionary()
    Set dict("properties") = properties
    Set dict("childIds") = childIds
    dict("backendDOMNodeId") = backendDOMNodeId

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set ignoredReasons = New Collection
    Set properties = New Collection
    Set childIds = New Collection
End Sub

