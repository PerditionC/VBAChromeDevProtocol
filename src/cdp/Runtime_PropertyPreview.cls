VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimePropertyPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PropertyPreview"
' Runtime.PropertyPreview
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Property name.
Public name AS string

' type : string
' Object type. Accessor means that the property itself is an accessor property.
Public fieldType AS string

' value : string
' Optional
' User-friendly property value string.
Public value AS string

' valuePreview : ObjectPreview
' Optional
' Nested value preview.
Public valuePreview AS cdpRuntimeObjectPreview

' subtype : string
' Optional
' Object subtype hint. Specified for `object` type values only.
Public subtype AS string

Public Function init(ByVal obj as Dictionary) As cdpRuntimePropertyPreview
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    Let fieldType = CStr(obj.Item("type"))
    If obj.Exists("value") Then Let value = CStr(obj.Item("value"))
    If obj.Exists("valuePreview") Then
        Set valuePreview = New cdpRuntimeObjectPreview
        valuePreview.init obj.Item("valuePreview")
    End If
    If obj.Exists("subtype") Then Let subtype = CStr(obj.Item("subtype"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("type") = fieldType
    if value <> vbNullString Then dict("value") = value
    If Not valuePreview Is Nothing Then Set dict("valuePreview") = valuePreview.asDictionary()
    if subtype <> vbNullString Then dict("subtype") = subtype

    Set asDictionary = dict
End Function

