VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCastSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sink"
' Cast.Sink
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' 
Public name AS string

' id : string
' 
Public id AS string

' session : string
' Optional
' Text describing the current session. Present only if there is an active
    '   session on the sink.
Public session AS string

Public Function init(ByVal obj as Dictionary) As cdpCastSink
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    Let id = CStr(obj.Item("id"))
    If obj.Exists("session") Then Let session = CStr(obj.Item("session"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("id") = id
    if session <> vbNullString Then dict("session") = session

    Set asDictionary = dict
End Function

