VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAccessibilityAXValueSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AXValueSource"
' Accessibility.AXValueSource
' A single source for a computed AX property.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : AXValueSourceType
' What type of source this is.
Public fieldType AS string

' value : AXValue
' Optional
' The value of this property source.
Public value AS cdpAccessibilityAXValue

' attribute : string
' Optional
' The name of the relevant attribute, if any.
Public fieldAttribute AS string

' attributeValue : AXValue
' Optional
' The value of the relevant attribute, if any.
Public attributeValue AS cdpAccessibilityAXValue

' superseded : boolean
' Optional
' Whether this source is superseded by a higher priority source.
Public superseded AS Variant ' boolean

' nativeSource : AXValueNativeSourceType
' Optional
' The native markup source for this value, e.g. a <label> element.
Public nativeSource AS string

' nativeSourceValue : AXValue
' Optional
' The value, such as a node or node list, of the native source.
Public nativeSourceValue AS cdpAccessibilityAXValue

' invalid : boolean
' Optional
' Whether the value for this property is invalid.
Public invalid AS Variant ' boolean

' invalidReason : string
' Optional
' Reason for the value being invalid, if it is.
Public invalidReason AS string

Public Function init(ByVal obj as Dictionary) As cdpAccessibilityAXValueSource
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let fieldType = CStr(obj.Item("type"))
    If obj.Exists("value") Then
        Set value = New cdpAccessibilityAXValue
        value.init obj.Item("value")
    End If
    If obj.Exists("attribute") Then Let fieldAttribute = CStr(obj.Item("attribute"))
    If obj.Exists("attributeValue") Then
        Set attributeValue = New cdpAccessibilityAXValue
        attributeValue.init obj.Item("attributeValue")
    End If
    If obj.Exists("superseded") Then Let superseded = CBool(obj.Item("superseded"))
    If obj.Exists("nativeSource") Then Let nativeSource = CStr(obj.Item("nativeSource"))
    If obj.Exists("nativeSourceValue") Then
        Set nativeSourceValue = New cdpAccessibilityAXValue
        nativeSourceValue.init obj.Item("nativeSourceValue")
    End If
    If obj.Exists("invalid") Then Let invalid = CBool(obj.Item("invalid"))
    If obj.Exists("invalidReason") Then Let invalidReason = CStr(obj.Item("invalidReason"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    If Not value Is Nothing Then Set dict("value") = value.asDictionary()
    if fieldAttribute <> vbNullString Then dict("attribute") = fieldAttribute
    If Not attributeValue Is Nothing Then Set dict("attributeValue") = attributeValue.asDictionary()
    if Not IsEmpty(superseded) Then dict("superseded") = superseded
    if nativeSource <> vbNullString Then dict("nativeSource") = nativeSource
    If Not nativeSourceValue Is Nothing Then Set dict("nativeSourceValue") = nativeSourceValue.asDictionary()
    if Not IsEmpty(invalid) Then dict("invalid") = invalid
    if invalidReason <> vbNullString Then dict("invalidReason") = invalidReason

    Set asDictionary = dict
End Function

