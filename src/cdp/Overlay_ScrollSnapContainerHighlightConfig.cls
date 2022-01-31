VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayScrollSnapContainerHi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScrollSnapContainerHighlightConfig"
' Overlay.ScrollSnapContainerHighlightConfig
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' snapportBorder : LineStyle
' Optional
' The style of the snapport border (default: transparent)
Public snapportBorder AS cdpOverlayLineStyle

' snapAreaBorder : LineStyle
' Optional
' The style of the snap area border (default: transparent)
Public snapAreaBorder AS cdpOverlayLineStyle

' scrollMarginColor : DOM.RGBA
' Optional
' The margin highlight fill color (default: transparent).
Public scrollMarginColor AS cdpDOMRGBA

' scrollPaddingColor : DOM.RGBA
' Optional
' The padding highlight fill color (default: transparent).
Public scrollPaddingColor AS cdpDOMRGBA

Public Function init(ByVal obj as Dictionary) As cdpOverlayScrollSnapContainerHi
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("snapportBorder") Then
        Set snapportBorder = New cdpOverlayLineStyle
        snapportBorder.init obj.Item("snapportBorder")
    End If
    If obj.Exists("snapAreaBorder") Then
        Set snapAreaBorder = New cdpOverlayLineStyle
        snapAreaBorder.init obj.Item("snapAreaBorder")
    End If
    If obj.Exists("scrollMarginColor") Then
        Set scrollMarginColor = New cdpDOMRGBA
        scrollMarginColor.init obj.Item("scrollMarginColor")
    End If
    If obj.Exists("scrollPaddingColor") Then
        Set scrollPaddingColor = New cdpDOMRGBA
        scrollPaddingColor.init obj.Item("scrollPaddingColor")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not snapportBorder Is Nothing Then Set dict("snapportBorder") = snapportBorder.asDictionary()
    If Not snapAreaBorder Is Nothing Then Set dict("snapAreaBorder") = snapAreaBorder.asDictionary()
    If Not scrollMarginColor Is Nothing Then Set dict("scrollMarginColor") = scrollMarginColor.asDictionary()
    If Not scrollPaddingColor Is Nothing Then Set dict("scrollPaddingColor") = scrollPaddingColor.asDictionary()

    Set asDictionary = dict
End Function

