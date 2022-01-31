VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayIsolationModeHighligh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "IsolationModeHighlightConfig"
' Overlay.IsolationModeHighlightConfig
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' resizerColor : DOM.RGBA
' Optional
' The fill color of the resizers (default: transparent).
Public resizerColor AS cdpDOMRGBA

' resizerHandleColor : DOM.RGBA
' Optional
' The fill color for resizer handles (default: transparent).
Public resizerHandleColor AS cdpDOMRGBA

' maskColor : DOM.RGBA
' Optional
' The fill color for the mask covering non-isolated elements (default: transparent).
Public maskColor AS cdpDOMRGBA

Public Function init(ByVal obj as Dictionary) As cdpOverlayIsolationModeHighligh
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("resizerColor") Then
        Set resizerColor = New cdpDOMRGBA
        resizerColor.init obj.Item("resizerColor")
    End If
    If obj.Exists("resizerHandleColor") Then
        Set resizerHandleColor = New cdpDOMRGBA
        resizerHandleColor.init obj.Item("resizerHandleColor")
    End If
    If obj.Exists("maskColor") Then
        Set maskColor = New cdpDOMRGBA
        maskColor.init obj.Item("maskColor")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not resizerColor Is Nothing Then Set dict("resizerColor") = resizerColor.asDictionary()
    If Not resizerHandleColor Is Nothing Then Set dict("resizerHandleColor") = resizerHandleColor.asDictionary()
    If Not maskColor Is Nothing Then Set dict("maskColor") = maskColor.asDictionary()

    Set asDictionary = dict
End Function

