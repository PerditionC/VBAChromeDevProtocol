VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Value"
' CSS.Value
' Data for a simple selector (these are delimited by commas in a selector list).
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' text : string
' Value text.
Public text AS string

' range : SourceRange
' Optional
' Value range in the underlying resource (if available).
Public range AS cdpCSSSourceRange

Public Function init(ByVal obj as Dictionary) As cdpCSSValue
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let text = CStr(obj.Item("text"))
    If obj.Exists("range") Then
        Set range = New cdpCSSSourceRange
        range.init obj.Item("range")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("text") = text
    If Not range Is Nothing Then Set dict("range") = range.asDictionary()

    Set asDictionary = dict
End Function

