VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSCSSRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CSSRule"
' CSS.CSSRule
' CSS rule representation.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' styleSheetId : StyleSheetId
' Optional
' The css style sheet identifier (absent for user agent stylesheet and user-specified
    '   stylesheet rules) this rule came from.
Public styleSheetId AS string

' selectorList : SelectorList
' Rule selector data.
Public selectorList AS cdpCSSSelectorList

' origin : StyleSheetOrigin
' Parent stylesheet's origin.
Public origin AS string

' style : CSSStyle
' Associated style declaration.
Public style AS cdpCSSCSSStyle

' media : array of cdpCSSCSSMedia
' Optional
' Media list array (for rules involving media queries). The array enumerates media queries
    '   starting with the innermost one, going outwards.
Public media AS Collection

' containerQueries : array of cdpCSSCSSContainerQuery
' Optional
' Container query list array (for rules involving container queries).
    '   The array enumerates container queries starting with the innermost one, going outwards.
Public containerQueries AS Collection

Public Function init(ByVal obj as Dictionary) As cdpCSSCSSRule
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    If obj.Exists("styleSheetId") Then Let styleSheetId = CStr(obj.Item("styleSheetId"))
    Set selectorList = New cdpCSSSelectorList
    selectorList.init obj.Item("selectorList")
    Let origin = CStr(obj.Item("origin"))
    Set style = New cdpCSSCSSStyle
    style.init obj.Item("style")
    If obj.Exists("media") Then
        For each v in obj.Item("media")
            Dim obj_media As cdpCSSCSSMedia: Set obj_media = New cdpCSSCSSMedia
            obj_media.init v
            media.Add obj_media
            Set obj_media = Nothing
        Next v
    End If
    If obj.Exists("containerQueries") Then
        For each v in obj.Item("containerQueries")
            Dim obj_containerQueries As cdpCSSCSSContainerQuery: Set obj_containerQueries = New cdpCSSCSSContainerQuery
            obj_containerQueries.init v
            containerQueries.Add obj_containerQueries
            Set obj_containerQueries = Nothing
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if styleSheetId <> vbNullString Then dict("styleSheetId") = styleSheetId
    Set dict("selectorList") = selectorList.asDictionary()
    dict("origin") = origin
    Set dict("style") = style.asDictionary()
    Set dict("media") = media
    Set dict("containerQueries") = containerQueries

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set media = New Collection
    Set containerQueries = New Collection
End Sub

