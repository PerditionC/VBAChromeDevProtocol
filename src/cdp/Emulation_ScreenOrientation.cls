VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpEmulationScreenOrientation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Emulation.ScreenOrientation
' Screen orientation.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : string
' Orientation type.
Public fieldType AS string

' angle : integer
' Orientation angle.
Public angle AS integer

Public Function init(ByVal obj as Dictionary) As cdpEmulationScreenOrientation

    Let fieldType = CStr(obj.Item("type"))
    Let angle = CInt(obj.Item("angle"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    dict("angle") = angle

    Set asDictionary = dict
End Function
