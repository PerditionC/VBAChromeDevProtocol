VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayHingeConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "HingeConfig"
' Overlay.HingeConfig
' Configuration for dual screen hinge
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' rect : DOM.Rect
' A rectangle represent hinge
Public rect AS cdpDOMRect

' contentColor : DOM.RGBA
' Optional
' The content box highlight fill color (default: a dark color).
Public contentColor AS cdpDOMRGBA

' outlineColor : DOM.RGBA
' Optional
' The content box highlight outline color (default: transparent).
Public outlineColor AS cdpDOMRGBA

Public Function init(ByVal obj as Dictionary) As cdpOverlayHingeConfig
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set rect = New cdpDOMRect
    rect.init obj.Item("rect")
    If obj.Exists("contentColor") Then
        Set contentColor = New cdpDOMRGBA
        contentColor.init obj.Item("contentColor")
    End If
    If obj.Exists("outlineColor") Then
        Set outlineColor = New cdpDOMRGBA
        outlineColor.init obj.Item("outlineColor")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("rect") = rect.asDictionary()
    If Not contentColor Is Nothing Then Set dict("contentColor") = contentColor.asDictionary()
    If Not outlineColor Is Nothing Then Set dict("outlineColor") = outlineColor.asDictionary()

    Set asDictionary = dict
End Function

