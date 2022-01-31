VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAudioAudioParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AudioParam"
' WebAudio.AudioParam
' Protocol object for AudioParam
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' paramId : GraphObjectId
' 
Public paramId AS string

' nodeId : GraphObjectId
' 
Public nodeId AS string

' contextId : GraphObjectId
' 
Public contextId AS string

' paramType : ParamType
' 
Public paramType AS string

' rate : AutomationRate
' 
Public rate AS string

' defaultValue : number
' 
Public defaultValue AS Double

' minValue : number
' 
Public minValue AS Double

' maxValue : number
' 
Public maxValue AS Double

Public Function init(ByVal obj as Dictionary) As cdpWebAudioAudioParam
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let paramId = CStr(obj.Item("paramId"))
    Let nodeId = CStr(obj.Item("nodeId"))
    Let contextId = CStr(obj.Item("contextId"))
    Let paramType = CStr(obj.Item("paramType"))
    Let rate = CStr(obj.Item("rate"))
    Let defaultValue = CDbl(obj.Item("defaultValue"))
    Let minValue = CDbl(obj.Item("minValue"))
    Let maxValue = CDbl(obj.Item("maxValue"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("paramId") = paramId
    dict("nodeId") = nodeId
    dict("contextId") = contextId
    dict("paramType") = paramType
    dict("rate") = rate
    dict("defaultValue") = defaultValue
    dict("minValue") = minValue
    dict("maxValue") = maxValue

    Set asDictionary = dict
End Function

