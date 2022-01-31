VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAudioAudioListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AudioListener"
' WebAudio.AudioListener
' Protocol object for AudioListener
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' listenerId : GraphObjectId
' 
Public listenerId AS string

' contextId : GraphObjectId
' 
Public contextId AS string

Public Function init(ByVal obj as Dictionary) As cdpWebAudioAudioListener
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let listenerId = CStr(obj.Item("listenerId"))
    Let contextId = CStr(obj.Item("contextId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("listenerId") = listenerId
    dict("contextId") = contextId

    Set asDictionary = dict
End Function

