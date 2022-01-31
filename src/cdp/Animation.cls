VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpAnimation"
' Animation [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Animation instance.
' Animation AS object

' AnimationEffect instance
' AnimationEffect AS object

' Keyframes Rule
' KeyframesRule AS object

' Keyframe Style
' KeyframeStyle AS object



Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' Disables animation domain notifications.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Animation.disable", params
End Sub

' Enables animation domain notifications.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Animation.enable", params
End Sub

' Returns the current time of the an animation.
' Returns: currentTime - Current time of the page.
Public Function getCurrentTime( _
    ByVal id AS string _
) AS Double
    ' id: string Id of animation.

    Dim params As New Dictionary
    params("id") = CStr(id)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Animation.getCurrentTime", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("currentTime") Then getCurrentTime = results("currentTime")
    End If
End Function

' Gets the playback rate of the document timeline.
' Returns: playbackRate - Playback rate for animations on page.
Public Function getPlaybackRate() AS Double
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Animation.getPlaybackRate", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("playbackRate") Then getPlaybackRate = results("playbackRate")
    End If
End Function

' Releases a set of animations to no longer be manipulated.
Public Sub releaseAnimations( _
    ByVal animations AS Collection _
)
    ' animations: array List of animation ids to seek.

    Dim params As New Dictionary
    Set params("animations") = animations

    cdp.InvokeMethod "Animation.releaseAnimations", params
End Sub

' Gets the remote object of the Animation.
' Returns: remoteObject - Corresponding remote object.
Public Function resolveAnimation( _
    ByVal animationId AS string _
) AS Dictionary
    ' animationId: string Animation id.

    Dim params As New Dictionary
    params("animationId") = CStr(animationId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Animation.resolveAnimation", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("remoteObject") Then Set resolveAnimation = results("remoteObject")
    End If
End Function

' Seek a set of animations to a particular time within each animation.
Public Sub seekAnimations( _
    ByVal animations AS Collection, _
    ByVal currentTime AS Double _
)
    ' animations: array List of animation ids to seek.
    ' currentTime: number Set the current time of each animation.

    Dim params As New Dictionary
    Set params("animations") = animations
    params("currentTime") = CDbl(currentTime)

    cdp.InvokeMethod "Animation.seekAnimations", params
End Sub

' Sets the paused state of a set of animations.
Public Sub setPaused( _
    ByVal animations AS Collection, _
    ByVal paused AS boolean _
)
    ' animations: array Animations to set the pause state of.
    ' paused: boolean Paused state to set to.

    Dim params As New Dictionary
    Set params("animations") = animations
    params("paused") = CBool(paused)

    cdp.InvokeMethod "Animation.setPaused", params
End Sub

' Sets the playback rate of the document timeline.
Public Sub setPlaybackRate( _
    ByVal playbackRate AS Double _
)
    ' playbackRate: number Playback rate for animations on page

    Dim params As New Dictionary
    params("playbackRate") = CDbl(playbackRate)

    cdp.InvokeMethod "Animation.setPlaybackRate", params
End Sub

' Sets the timing of an animation node.
Public Sub setTiming( _
    ByVal animationId AS string, _
    ByVal duration AS Double, _
    ByVal delay AS Double _
)
    ' animationId: string Animation id.
    ' duration: number Duration of the animation.
    ' delay: number Delay of the animation.

    Dim params As New Dictionary
    params("animationId") = CStr(animationId)
    params("duration") = CDbl(duration)
    params("delay") = CDbl(delay)

    cdp.InvokeMethod "Animation.setTiming", params
End Sub

