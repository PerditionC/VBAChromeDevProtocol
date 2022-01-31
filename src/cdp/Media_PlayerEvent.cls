VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpMediaPlayerEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PlayerEvent"
' Media.PlayerEvent
' Corresponds to kMediaEventTriggered
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' timestamp : Timestamp
' 
Public timestamp AS Double

' value : string
' 
Public value AS string

Public Function init(ByVal obj as Dictionary) As cdpMediaPlayerEvent
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let timestamp = CDbl(obj.Item("timestamp"))
    Let value = CStr(obj.Item("value"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("timestamp") = timestamp
    dict("value") = value

    Set asDictionary = dict
End Function

