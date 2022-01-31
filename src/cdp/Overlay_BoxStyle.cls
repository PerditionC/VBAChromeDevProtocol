VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayBoxStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "BoxStyle"
' Overlay.BoxStyle
' Style information for drawing a box.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' fillColor : DOM.RGBA
' Optional
' The background color for the box (default: transparent)
Public fillColor AS cdpDOMRGBA

' hatchColor : DOM.RGBA
' Optional
' The hatching color for the box (default: transparent)
Public hatchColor AS cdpDOMRGBA

Public Function init(ByVal obj as Dictionary) As cdpOverlayBoxStyle
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("fillColor") Then
        Set fillColor = New cdpDOMRGBA
        fillColor.init obj.Item("fillColor")
    End If
    If obj.Exists("hatchColor") Then
        Set hatchColor = New cdpDOMRGBA
        hatchColor.init obj.Item("hatchColor")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not fillColor Is Nothing Then Set dict("fillColor") = fillColor.asDictionary()
    If Not hatchColor Is Nothing Then Set dict("hatchColor") = hatchColor.asDictionary()

    Set asDictionary = dict
End Function

