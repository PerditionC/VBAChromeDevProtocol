VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSFontVariationAxis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FontVariationAxis"
' CSS.FontVariationAxis
' Information about font variation axes for variable fonts
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' tag : string
' The font-variation-setting tag (a.k.a. "axis tag").
Public tag AS string

' name : string
' Human-readable variation name in the default language (normally, "en").
Public name AS string

' minValue : number
' The minimum value (inclusive) the font supports for this tag.
Public minValue AS Double

' maxValue : number
' The maximum value (inclusive) the font supports for this tag.
Public maxValue AS Double

' defaultValue : number
' The default value.
Public defaultValue AS Double

Public Function init(ByVal obj as Dictionary) As cdpCSSFontVariationAxis
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let tag = CStr(obj.Item("tag"))
    Let name = CStr(obj.Item("name"))
    Let minValue = CDbl(obj.Item("minValue"))
    Let maxValue = CDbl(obj.Item("maxValue"))
    Let defaultValue = CDbl(obj.Item("defaultValue"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("tag") = tag
    dict("name") = name
    dict("minValue") = minValue
    dict("maxValue") = maxValue
    dict("defaultValue") = defaultValue

    Set asDictionary = dict
End Function

