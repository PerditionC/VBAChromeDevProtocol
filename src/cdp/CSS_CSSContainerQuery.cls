VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSCSSContainerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CSSContainerQuery"
' CSS.CSSContainerQuery
' CSS container query rule descriptor.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' text : string
' Container query text.
Public text AS string

' range : SourceRange
' Optional
' The associated rule header range in the enclosing stylesheet (if
    '   available).
Public range AS cdpCSSSourceRange

' styleSheetId : StyleSheetId
' Optional
' Identifier of the stylesheet containing this object (if exists).
Public styleSheetId AS string

' name : string
' Optional
' Optional name for the container.
Public name AS string

Public Function init(ByVal obj as Dictionary) As cdpCSSCSSContainerQuery
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let text = CStr(obj.Item("text"))
    If obj.Exists("range") Then
        Set range = New cdpCSSSourceRange
        range.init obj.Item("range")
    End If
    If obj.Exists("styleSheetId") Then Let styleSheetId = CStr(obj.Item("styleSheetId"))
    If obj.Exists("name") Then Let name = CStr(obj.Item("name"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("text") = text
    If Not range Is Nothing Then Set dict("range") = range.asDictionary()
    if styleSheetId <> vbNullString Then dict("styleSheetId") = styleSheetId
    if name <> vbNullString Then dict("name") = name

    Set asDictionary = dict
End Function

