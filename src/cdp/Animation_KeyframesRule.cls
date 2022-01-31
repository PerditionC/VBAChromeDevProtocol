VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAnimationKeyframesRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "KeyframesRule"
' Animation.KeyframesRule
' Keyframes Rule
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Optional
' CSS keyframed animation's name.
Public name AS string

' keyframes : array of cdpAnimationKeyframeStyle
' List of animation keyframes.
Public keyframes AS Collection

Public Function init(ByVal obj as Dictionary) As cdpAnimationKeyframesRule
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    If obj.Exists("name") Then Let name = CStr(obj.Item("name"))
    For each v in obj.Item("keyframes")
        Dim obj_keyframes As cdpAnimationKeyframeStyle: Set obj_keyframes = New cdpAnimationKeyframeStyle
        obj_keyframes.init v
        keyframes.Add obj_keyframes
        Set obj_keyframes = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if name <> vbNullString Then dict("name") = name
    Set dict("keyframes") = keyframes

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set keyframes = New Collection
End Sub

