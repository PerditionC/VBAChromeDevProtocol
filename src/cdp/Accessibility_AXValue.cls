VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAccessibilityAXValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AXValue"
' Accessibility.AXValue
' A single computed AX property.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : AXValueType
' The type of this value.
Public fieldType AS string

' value : any
' Optional
' The computed value of this property.
Public value AS Variant

' relatedNodes : array of cdpAccessibilityAXRelatedNode
' Optional
' One or more related nodes, if applicable.
Public relatedNodes AS Collection

' sources : array of cdpAccessibilityAXValueSource
' Optional
' The sources which contributed to the computation of this property.
Public sources AS Collection

Public Function init(ByVal obj as Dictionary) As cdpAccessibilityAXValue
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let fieldType = CStr(obj.Item("type"))
    If obj.Exists("value") Then
        If IsObject(obj.Item("value")) Then
            Set value = obj.Item("value")
        Else
            Let value = obj.Item("value")
        End If
    End If
    If obj.Exists("relatedNodes") Then
        For each v in obj.Item("relatedNodes")
            Dim obj_relatedNodes As cdpAccessibilityAXRelatedNode: Set obj_relatedNodes = New cdpAccessibilityAXRelatedNode
            obj_relatedNodes.init v
            relatedNodes.Add obj_relatedNodes
            Set obj_relatedNodes = Nothing
        Next v
    End If
    If obj.Exists("sources") Then
        For each v in obj.Item("sources")
            Dim obj_sources As cdpAccessibilityAXValueSource: Set obj_sources = New cdpAccessibilityAXValueSource
            obj_sources.init v
            sources.Add obj_sources
            Set obj_sources = Nothing
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    dict("value") = value
    Set dict("relatedNodes") = relatedNodes
    Set dict("sources") = sources

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set relatedNodes = New Collection
    Set sources = New Collection
End Sub

