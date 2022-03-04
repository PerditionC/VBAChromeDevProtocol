VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAnimationAnimationEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AnimationEffect"
' Animation.AnimationEffect
' AnimationEffect instance
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' delay : number
' `AnimationEffect`'s delay.
Public delay AS Double

' endDelay : number
' `AnimationEffect`'s end delay.
Public endDelay AS Double

' iterationStart : number
' `AnimationEffect`'s iteration start.
Public iterationStart AS Double

' iterations : number
' `AnimationEffect`'s iterations.
Public iterations AS Double

' duration : number
' `AnimationEffect`'s iteration duration.
Public duration AS Double

' direction : string
' `AnimationEffect`'s playback direction.
Public direction AS string

' fill : string
' `AnimationEffect`'s fill mode.
Public fill AS string

' backendNodeId : DOM.BackendNodeId
' Optional
' `AnimationEffect`'s target node.
Public backendNodeId AS Long

' keyframesRule : KeyframesRule
' Optional
' `AnimationEffect`'s keyframes.
Public keyframesRule AS cdpAnimationKeyframesRule

' easing : string
' `AnimationEffect`'s timing function.
Public easing AS string

Public Function init(ByVal obj as Dictionary) As cdpAnimationAnimationEffect
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let delay = CDbl(obj.Item("delay"))
    Let endDelay = CDbl(obj.Item("endDelay"))
    Let iterationStart = CDbl(obj.Item("iterationStart"))
    Let iterations = CDbl(obj.Item("iterations"))
    Let duration = CDbl(obj.Item("duration"))
    Let direction = CStr(obj.Item("direction"))
    Let fill = CStr(obj.Item("fill"))
    If obj.Exists("backendNodeId") Then Let backendNodeId = CLng(obj.Item("backendNodeId"))
    If obj.Exists("keyframesRule") Then
        Set keyframesRule = New cdpAnimationKeyframesRule
        keyframesRule.init obj.Item("keyframesRule")
    End If
    Let easing = CStr(obj.Item("easing"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("delay") = delay
    dict("endDelay") = endDelay
    dict("iterationStart") = iterationStart
    dict("iterations") = iterations
    dict("duration") = duration
    dict("direction") = direction
    dict("fill") = fill
    dict("backendNodeId") = backendNodeId
    If Not keyframesRule Is Nothing Then Set dict("keyframesRule") = keyframesRule.asDictionary()
    dict("easing") = easing

    Set asDictionary = dict
End Function

