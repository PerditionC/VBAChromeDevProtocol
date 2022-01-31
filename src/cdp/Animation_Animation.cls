VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAnimationAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Animation"
' Animation.Animation
' Animation instance.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' id : string
' `Animation`'s id.
Public id AS string

' name : string
' `Animation`'s name.
Public name AS string

' pausedState : boolean
' `Animation`'s internal paused state.
Public pausedState AS boolean

' playState : string
' `Animation`'s play state.
Public playState AS string

' playbackRate : number
' `Animation`'s playback rate.
Public playbackRate AS Double

' startTime : number
' `Animation`'s start time.
Public startTime AS Double

' currentTime : number
' `Animation`'s current time.
Public currentTime AS Double

' type : string
' Animation type of `Animation`.
Public fieldType AS string

' source : AnimationEffect
' Optional
' `Animation`'s source animation node.
Public source AS cdpAnimationAnimationEffect

' cssId : string
' Optional
' A unique ID for `Animation` representing the sources that triggered this CSS
    '   animation/transition.
Public cssId AS string

Public Function init(ByVal obj as Dictionary) As cdpAnimationAnimation
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let id = CStr(obj.Item("id"))
    Let name = CStr(obj.Item("name"))
    Let pausedState = CBool(obj.Item("pausedState"))
    Let playState = CStr(obj.Item("playState"))
    Let playbackRate = CDbl(obj.Item("playbackRate"))
    Let startTime = CDbl(obj.Item("startTime"))
    Let currentTime = CDbl(obj.Item("currentTime"))
    Let fieldType = CStr(obj.Item("type"))
    If obj.Exists("source") Then
        Set source = New cdpAnimationAnimationEffect
        source.init obj.Item("source")
    End If
    If obj.Exists("cssId") Then Let cssId = CStr(obj.Item("cssId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("id") = id
    dict("name") = name
    dict("pausedState") = pausedState
    dict("playState") = playState
    dict("playbackRate") = playbackRate
    dict("startTime") = startTime
    dict("currentTime") = currentTime
    dict("type") = fieldType
    If Not source Is Nothing Then Set dict("source") = source.asDictionary()
    if cssId <> vbNullString Then dict("cssId") = cssId

    Set asDictionary = dict
End Function

