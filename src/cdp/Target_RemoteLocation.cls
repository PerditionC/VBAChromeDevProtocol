VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpTargetRemoteLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Target.RemoteLocation
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' host : string
' 
Public host AS string

' port : integer
' 
Public port AS integer

Public Function init(ByVal obj as Dictionary) As cdpTargetRemoteLocation

    Let host = CStr(obj.Item("host"))
    Let port = CInt(obj.Item("port"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("host") = host
    dict("port") = port

    Set asDictionary = dict
End Function

