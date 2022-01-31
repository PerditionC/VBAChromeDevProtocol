VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSCSSStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CSSStyle"
' CSS.CSSStyle
' CSS style representation.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' styleSheetId : StyleSheetId
' Optional
' The css style sheet identifier (absent for user agent stylesheet and user-specified
    '   stylesheet rules) this rule came from.
Public styleSheetId AS string

' cssProperties : array of cdpCSSCSSProperty
' CSS properties in the style.
Public cssProperties AS Collection

' shorthandEntries : array of cdpCSSShorthandEntry
' Computed values for all shorthands found in the style.
Public shorthandEntries AS Collection

' cssText : string
' Optional
' Style declaration text (if available).
Public cssText AS string

' range : SourceRange
' Optional
' Style declaration range in the enclosing stylesheet (if available).
Public range AS cdpCSSSourceRange

Public Function init(ByVal obj as Dictionary) As cdpCSSCSSStyle
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    If obj.Exists("styleSheetId") Then Let styleSheetId = CStr(obj.Item("styleSheetId"))
    For each v in obj.Item("cssProperties")
        Dim obj_cssProperties As cdpCSSCSSProperty: Set obj_cssProperties = New cdpCSSCSSProperty
        obj_cssProperties.init v
        cssProperties.Add obj_cssProperties
        Set obj_cssProperties = Nothing
    Next v
    For each v in obj.Item("shorthandEntries")
        Dim obj_shorthandEntries As cdpCSSShorthandEntry: Set obj_shorthandEntries = New cdpCSSShorthandEntry
        obj_shorthandEntries.init v
        shorthandEntries.Add obj_shorthandEntries
        Set obj_shorthandEntries = Nothing
    Next v
    If obj.Exists("cssText") Then Let cssText = CStr(obj.Item("cssText"))
    If obj.Exists("range") Then
        Set range = New cdpCSSSourceRange
        range.init obj.Item("range")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if styleSheetId <> vbNullString Then dict("styleSheetId") = styleSheetId
    Set dict("cssProperties") = cssProperties
    Set dict("shorthandEntries") = shorthandEntries
    if cssText <> vbNullString Then dict("cssText") = cssText
    If Not range Is Nothing Then Set dict("range") = range.asDictionary()

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set cssProperties = New Collection
    Set shorthandEntries = New Collection
End Sub

