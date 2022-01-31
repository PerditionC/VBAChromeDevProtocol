VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSCSSProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CSSProperty"
' CSS.CSSProperty
' CSS property declaration data.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' The property name.
Public name AS string

' value : string
' The property value.
Public value AS string

' important : boolean
' Optional
' Whether the property has "!important" annotation (implies `false` if absent).
Public important AS Variant ' boolean

' implicit : boolean
' Optional
' Whether the property is implicit (implies `false` if absent).
Public implicit AS Variant ' boolean

' text : string
' Optional
' The full property text as specified in the style.
Public text AS string

' parsedOk : boolean
' Optional
' Whether the property is understood by the browser (implies `true` if absent).
Public parsedOk AS Variant ' boolean

' disabled : boolean
' Optional
' Whether the property is disabled by the user (present for source-based properties only).
Public disabled AS Variant ' boolean

' range : SourceRange
' Optional
' The entire property range in the enclosing style declaration (if available).
Public range AS cdpCSSSourceRange

Public Function init(ByVal obj as Dictionary) As cdpCSSCSSProperty
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    Let value = CStr(obj.Item("value"))
    If obj.Exists("important") Then Let important = CBool(obj.Item("important"))
    If obj.Exists("implicit") Then Let implicit = CBool(obj.Item("implicit"))
    If obj.Exists("text") Then Let text = CStr(obj.Item("text"))
    If obj.Exists("parsedOk") Then Let parsedOk = CBool(obj.Item("parsedOk"))
    If obj.Exists("disabled") Then Let disabled = CBool(obj.Item("disabled"))
    If obj.Exists("range") Then
        Set range = New cdpCSSSourceRange
        range.init obj.Item("range")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("value") = value
    if Not IsEmpty(important) Then dict("important") = important
    if Not IsEmpty(implicit) Then dict("implicit") = implicit
    if text <> vbNullString Then dict("text") = text
    if Not IsEmpty(parsedOk) Then dict("parsedOk") = parsedOk
    if Not IsEmpty(disabled) Then dict("disabled") = disabled
    If Not range Is Nothing Then Set dict("range") = range.asDictionary()

    Set asDictionary = dict
End Function

