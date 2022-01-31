VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpEmulationMediaFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Emulation.MediaFeature
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' 
Public name AS string

' value : string
' 
Public value AS string

Public Function init(ByVal obj as Dictionary) As cdpEmulationMediaFeature

    Let name = CStr(obj.Item("name"))
    Let value = CStr(obj.Item("value"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("value") = value

    Set asDictionary = dict
End Function

