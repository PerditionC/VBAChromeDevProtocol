VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSMediaQueryExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "MediaQueryExpression"
' CSS.MediaQueryExpression
' Media query expression descriptor.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' value : number
' Media query expression value.
Public value AS Double

' unit : string
' Media query expression units.
Public unit AS string

' feature : string
' Media query expression feature.
Public feature AS string

' valueRange : SourceRange
' Optional
' The associated range of the value text in the enclosing stylesheet (if available).
Public valueRange AS cdpCSSSourceRange

' computedLength : number
' Optional
' Computed length of media query expression (if applicable).
Public computedLength AS Double

Public Function init(ByVal obj as Dictionary) As cdpCSSMediaQueryExpression
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let value = CDbl(obj.Item("value"))
    Let unit = CStr(obj.Item("unit"))
    Let feature = CStr(obj.Item("feature"))
    If obj.Exists("valueRange") Then
        Set valueRange = New cdpCSSSourceRange
        valueRange.init obj.Item("valueRange")
    End If
    If obj.Exists("computedLength") Then Let computedLength = CDbl(obj.Item("computedLength"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("value") = value
    dict("unit") = unit
    dict("feature") = feature
    If Not valueRange Is Nothing Then Set dict("valueRange") = valueRange.asDictionary()
    dict("computedLength") = computedLength

    Set asDictionary = dict
End Function

