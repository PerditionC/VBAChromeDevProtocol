VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAudioBaseAudioContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "BaseAudioContext"
' WebAudio.BaseAudioContext
' Protocol object for BaseAudioContext
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' contextId : GraphObjectId
' 
Public contextId AS string

' contextType : ContextType
' 
Public contextType AS string

' contextState : ContextState
' 
Public contextState AS string

' realtimeData : ContextRealtimeData
' Optional
' 
Public realtimeData AS cdpWebAudioContextRealtimeData

' callbackBufferSize : number
' Platform-dependent callback buffer size.
Public callbackBufferSize AS Double

' maxOutputChannelCount : number
' Number of output channels supported by audio hardware in use.
Public maxOutputChannelCount AS Double

' sampleRate : number
' Context sample rate.
Public sampleRate AS Double

Public Function init(ByVal obj as Dictionary) As cdpWebAudioBaseAudioContext
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let contextId = CStr(obj.Item("contextId"))
    Let contextType = CStr(obj.Item("contextType"))
    Let contextState = CStr(obj.Item("contextState"))
    If obj.Exists("realtimeData") Then
        Set realtimeData = New cdpWebAudioContextRealtimeData
        realtimeData.init obj.Item("realtimeData")
    End If
    Let callbackBufferSize = CDbl(obj.Item("callbackBufferSize"))
    Let maxOutputChannelCount = CDbl(obj.Item("maxOutputChannelCount"))
    Let sampleRate = CDbl(obj.Item("sampleRate"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("contextId") = contextId
    dict("contextType") = contextType
    dict("contextState") = contextState
    If Not realtimeData Is Nothing Then Set dict("realtimeData") = realtimeData.asDictionary()
    dict("callbackBufferSize") = callbackBufferSize
    dict("maxOutputChannelCount") = maxOutputChannelCount
    dict("sampleRate") = sampleRate

    Set asDictionary = dict
End Function

