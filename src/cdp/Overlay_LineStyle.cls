VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayLineStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "LineStyle"
' Overlay.LineStyle
' Style information for drawing a line.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' color : DOM.RGBA
' Optional
' The color of the line (default: transparent)
Public color AS cdpDOMRGBA

' pattern : string
' Optional
' The line pattern (default: solid)
Public pattern AS string

Public Function init(ByVal obj as Dictionary) As cdpOverlayLineStyle
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("color") Then
        Set color = New cdpDOMRGBA
        color.init obj.Item("color")
    End If
    If obj.Exists("pattern") Then Let pattern = CStr(obj.Item("pattern"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not color Is Nothing Then Set dict("color") = color.asDictionary()
    if pattern <> vbNullString Then dict("pattern") = pattern

    Set asDictionary = dict
End Function

