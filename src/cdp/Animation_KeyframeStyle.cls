VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAnimationKeyframeStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "KeyframeStyle"
' Animation.KeyframeStyle
' Keyframe Style
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' offset : string
' Keyframe's time offset.
Public offset AS string

' easing : string
' `AnimationEffect`'s timing function.
Public easing AS string

Public Function init(ByVal obj as Dictionary) As cdpAnimationKeyframeStyle
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let offset = CStr(obj.Item("offset"))
    Let easing = CStr(obj.Item("easing"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("offset") = offset
    dict("easing") = easing

    Set asDictionary = dict
End Function

