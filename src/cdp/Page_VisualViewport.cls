VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageVisualViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VisualViewport"
' Page.VisualViewport
' Visual viewport position, dimensions, and scale.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' offsetX : number
' Horizontal offset relative to the layout viewport (CSS pixels).
Public offsetX AS Double

' offsetY : number
' Vertical offset relative to the layout viewport (CSS pixels).
Public offsetY AS Double

' pageX : number
' Horizontal offset relative to the document (CSS pixels).
Public pageX AS Double

' pageY : number
' Vertical offset relative to the document (CSS pixels).
Public pageY AS Double

' clientWidth : number
' Width (CSS pixels), excludes scrollbar if present.
Public clientWidth AS Double

' clientHeight : number
' Height (CSS pixels), excludes scrollbar if present.
Public clientHeight AS Double

' scale : number
' Scale relative to the ideal viewport (size at width=device-width).
Public fieldScale AS Double

' zoom : number
' Optional
' Page zoom factor (CSS to device independent pixels ratio).
Public zoom AS Double

Public Function init(ByVal obj as Dictionary) As cdpPageVisualViewport
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let offsetX = CDbl(obj.Item("offsetX"))
    Let offsetY = CDbl(obj.Item("offsetY"))
    Let pageX = CDbl(obj.Item("pageX"))
    Let pageY = CDbl(obj.Item("pageY"))
    Let clientWidth = CDbl(obj.Item("clientWidth"))
    Let clientHeight = CDbl(obj.Item("clientHeight"))
    Let fieldScale = CDbl(obj.Item("scale"))
    If obj.Exists("zoom") Then Let zoom = CDbl(obj.Item("zoom"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("offsetX") = offsetX
    dict("offsetY") = offsetY
    dict("pageX") = pageX
    dict("pageY") = pageY
    dict("clientWidth") = clientWidth
    dict("clientHeight") = clientHeight
    dict("scale") = fieldScale
    dict("zoom") = zoom

    Set asDictionary = dict
End Function

