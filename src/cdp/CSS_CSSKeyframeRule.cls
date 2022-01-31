VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSCSSKeyframeRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CSSKeyframeRule"
' CSS.CSSKeyframeRule
' CSS keyframe rule representation.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' styleSheetId : StyleSheetId
' Optional
' The css style sheet identifier (absent for user agent stylesheet and user-specified
    '   stylesheet rules) this rule came from.
Public styleSheetId AS string

' origin : StyleSheetOrigin
' Parent stylesheet's origin.
Public origin AS string

' keyText : Value
' Associated key text.
Public keyText AS cdpCSSValue

' style : CSSStyle
' Associated style declaration.
Public style AS cdpCSSCSSStyle

Public Function init(ByVal obj as Dictionary) As cdpCSSCSSKeyframeRule
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("styleSheetId") Then Let styleSheetId = CStr(obj.Item("styleSheetId"))
    Let origin = CStr(obj.Item("origin"))
    Set keyText = New cdpCSSValue
    keyText.init obj.Item("keyText")
    Set style = New cdpCSSCSSStyle
    style.init obj.Item("style")

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if styleSheetId <> vbNullString Then dict("styleSheetId") = styleSheetId
    dict("origin") = origin
    Set dict("keyText") = keyText.asDictionary()
    Set dict("style") = style.asDictionary()

    Set asDictionary = dict
End Function

