VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAudioContextRealtimeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ContextRealtimeData"
' WebAudio.ContextRealtimeData
' Fields in AudioContext that change in real-time.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' currentTime : number
' The current context time in second in BaseAudioContext.
Public currentTime AS Double

' renderCapacity : number
' The time spent on rendering graph divided by render quantum duration,
    '   and multiplied by 100. 100 means the audio renderer reached the full
    '   capacity and glitch may occur.
Public renderCapacity AS Double

' callbackIntervalMean : number
' A running mean of callback interval.
Public callbackIntervalMean AS Double

' callbackIntervalVariance : number
' A running variance of callback interval.
Public callbackIntervalVariance AS Double

Public Function init(ByVal obj as Dictionary) As cdpWebAudioContextRealtimeData
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let currentTime = CDbl(obj.Item("currentTime"))
    Let renderCapacity = CDbl(obj.Item("renderCapacity"))
    Let callbackIntervalMean = CDbl(obj.Item("callbackIntervalMean"))
    Let callbackIntervalVariance = CDbl(obj.Item("callbackIntervalVariance"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("currentTime") = currentTime
    dict("renderCapacity") = renderCapacity
    dict("callbackIntervalMean") = callbackIntervalMean
    dict("callbackIntervalVariance") = callbackIntervalVariance

    Set asDictionary = dict
End Function

