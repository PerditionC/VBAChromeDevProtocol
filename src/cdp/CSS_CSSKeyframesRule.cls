VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSCSSKeyframesRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CSSKeyframesRule"
' CSS.CSSKeyframesRule
' CSS keyframes rule representation.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' animationName : Value
' Animation name.
Public animationName AS cdpCSSValue

' keyframes : array of cdpCSSCSSKeyframeRule
' List of keyframes.
Public keyframes AS Collection

Public Function init(ByVal obj as Dictionary) As cdpCSSCSSKeyframesRule
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Set animationName = New cdpCSSValue
    animationName.init obj.Item("animationName")
    For each v in obj.Item("keyframes")
        Dim obj_keyframes As cdpCSSCSSKeyframeRule: Set obj_keyframes = New cdpCSSCSSKeyframeRule
        obj_keyframes.init v
        keyframes.Add obj_keyframes
        Set obj_keyframes = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("animationName") = animationName.asDictionary()
    Set dict("keyframes") = keyframes

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set keyframes = New Collection
End Sub

