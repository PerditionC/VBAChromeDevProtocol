VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBackgroundServiceEventMetada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "EventMetadata"
' BackgroundService.EventMetadata
' A key-value pair for additional event information to pass along.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' key : string
' 
Public key AS string

' value : string
' 
Public value AS string

Public Function init(ByVal obj as Dictionary) As cdpBackgroundServiceEventMetada
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let key = CStr(obj.Item("key"))
    Let value = CStr(obj.Item("value"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("key") = key
    dict("value") = value

    Set asDictionary = dict
End Function

