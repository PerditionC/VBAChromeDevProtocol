VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSCSSMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CSSMedia"
' CSS.CSSMedia
' CSS media rule descriptor.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' text : string
' Media query text.
Public text AS string

' source : string
' Source of the media query: "mediaRule" if specified by a @media rule, "importRule" if
    '   specified by an @import rule, "linkedSheet" if specified by a "media" attribute in a linked
    '   stylesheet's LINK tag, "inlineSheet" if specified by a "media" attribute in an inline
    '   stylesheet's STYLE tag.
Public source AS string

' sourceURL : string
' Optional
' URL of the document containing the media query description.
Public sourceURL AS string

' range : SourceRange
' Optional
' The associated rule (@media or @import) header range in the enclosing stylesheet (if
    '   available).
Public range AS cdpCSSSourceRange

' styleSheetId : StyleSheetId
' Optional
' Identifier of the stylesheet containing this object (if exists).
Public styleSheetId AS string

' mediaList : array of cdpCSSMediaQuery
' Optional
' Array of media queries.
Public mediaList AS Collection

Public Function init(ByVal obj as Dictionary) As cdpCSSCSSMedia
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let text = CStr(obj.Item("text"))
    Let source = CStr(obj.Item("source"))
    If obj.Exists("sourceURL") Then Let sourceURL = CStr(obj.Item("sourceURL"))
    If obj.Exists("range") Then
        Set range = New cdpCSSSourceRange
        range.init obj.Item("range")
    End If
    If obj.Exists("styleSheetId") Then Let styleSheetId = CStr(obj.Item("styleSheetId"))
    If obj.Exists("mediaList") Then
        For each v in obj.Item("mediaList")
            Dim obj_mediaList As cdpCSSMediaQuery: Set obj_mediaList = New cdpCSSMediaQuery
            obj_mediaList.init v
            mediaList.Add obj_mediaList
            Set obj_mediaList = Nothing
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("text") = text
    dict("source") = source
    if sourceURL <> vbNullString Then dict("sourceURL") = sourceURL
    If Not range Is Nothing Then Set dict("range") = range.asDictionary()
    if styleSheetId <> vbNullString Then dict("styleSheetId") = styleSheetId
    Set dict("mediaList") = mediaList

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set mediaList = New Collection
End Sub

