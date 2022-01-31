VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAudioAudioNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AudioNode"
' WebAudio.AudioNode
' Protocol object for AudioNode
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' nodeId : GraphObjectId
' 
Public nodeId AS string

' contextId : GraphObjectId
' 
Public contextId AS string

' nodeType : NodeType
' 
Public nodeType AS string

' numberOfInputs : number
' 
Public numberOfInputs AS Double

' numberOfOutputs : number
' 
Public numberOfOutputs AS Double

' channelCount : number
' 
Public channelCount AS Double

' channelCountMode : ChannelCountMode
' 
Public channelCountMode AS string

' channelInterpretation : ChannelInterpretation
' 
Public channelInterpretation AS string

Public Function init(ByVal obj as Dictionary) As cdpWebAudioAudioNode
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let nodeId = CStr(obj.Item("nodeId"))
    Let contextId = CStr(obj.Item("contextId"))
    Let nodeType = CStr(obj.Item("nodeType"))
    Let numberOfInputs = CDbl(obj.Item("numberOfInputs"))
    Let numberOfOutputs = CDbl(obj.Item("numberOfOutputs"))
    Let channelCount = CDbl(obj.Item("channelCount"))
    Let channelCountMode = CStr(obj.Item("channelCountMode"))
    Let channelInterpretation = CStr(obj.Item("channelInterpretation"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("nodeId") = nodeId
    dict("contextId") = contextId
    dict("nodeType") = nodeType
    dict("numberOfInputs") = numberOfInputs
    dict("numberOfOutputs") = numberOfOutputs
    dict("channelCount") = channelCount
    dict("channelCountMode") = channelCountMode
    dict("channelInterpretation") = channelInterpretation

    Set asDictionary = dict
End Function

