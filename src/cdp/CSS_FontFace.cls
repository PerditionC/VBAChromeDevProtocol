VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSFontFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FontFace"
' CSS.FontFace
' Properties of a web font: https://www.w3.org/TR/2008/REC-CSS2-20080411/fonts.html#font-descriptions
'   and additional information such as platformFontFamily and fontVariationAxes.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' fontFamily : string
' The font-family.
Public fontFamily AS string

' fontStyle : string
' The font-style.
Public fontStyle AS string

' fontVariant : string
' The font-variant.
Public fontVariant AS string

' fontWeight : string
' The font-weight.
Public fontWeight AS string

' fontStretch : string
' The font-stretch.
Public fontStretch AS string

' unicodeRange : string
' The unicode-range.
Public unicodeRange AS string

' src : string
' The src.
Public src AS string

' platformFontFamily : string
' The resolved platform font family
Public platformFontFamily AS string

' fontVariationAxes : array of cdpCSSFontVariationAxis
' Optional
' Available variation settings (a.k.a. "axes").
Public fontVariationAxes AS Collection

Public Function init(ByVal obj as Dictionary) As cdpCSSFontFace
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let fontFamily = CStr(obj.Item("fontFamily"))
    Let fontStyle = CStr(obj.Item("fontStyle"))
    Let fontVariant = CStr(obj.Item("fontVariant"))
    Let fontWeight = CStr(obj.Item("fontWeight"))
    Let fontStretch = CStr(obj.Item("fontStretch"))
    Let unicodeRange = CStr(obj.Item("unicodeRange"))
    Let src = CStr(obj.Item("src"))
    Let platformFontFamily = CStr(obj.Item("platformFontFamily"))
    If obj.Exists("fontVariationAxes") Then
        For each v in obj.Item("fontVariationAxes")
            Dim obj_fontVariationAxes As cdpCSSFontVariationAxis: Set obj_fontVariationAxes = New cdpCSSFontVariationAxis
            obj_fontVariationAxes.init v
            fontVariationAxes.Add obj_fontVariationAxes
            Set obj_fontVariationAxes = Nothing
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("fontFamily") = fontFamily
    dict("fontStyle") = fontStyle
    dict("fontVariant") = fontVariant
    dict("fontWeight") = fontWeight
    dict("fontStretch") = fontStretch
    dict("unicodeRange") = unicodeRange
    dict("src") = src
    dict("platformFontFamily") = platformFontFamily
    Set dict("fontVariationAxes") = fontVariationAxes

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set fontVariationAxes = New Collection
End Sub

