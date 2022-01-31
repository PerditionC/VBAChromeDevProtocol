VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSPlatformFontUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PlatformFontUsage"
' CSS.PlatformFontUsage
' Information about amount of glyphs that were rendered with given font.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' familyName : string
' Font's family name reported by platform.
Public familyName AS string

' isCustomFont : boolean
' Indicates if the font was downloaded or resolved locally.
Public isCustomFont AS boolean

' glyphCount : number
' Amount of glyphs that were rendered with this font.
Public glyphCount AS Double

Public Function init(ByVal obj as Dictionary) As cdpCSSPlatformFontUsage
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let familyName = CStr(obj.Item("familyName"))
    Let isCustomFont = CBool(obj.Item("isCustomFont"))
    Let glyphCount = CDbl(obj.Item("glyphCount"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("familyName") = familyName
    dict("isCustomFont") = isCustomFont
    dict("glyphCount") = glyphCount

    Set asDictionary = dict
End Function

