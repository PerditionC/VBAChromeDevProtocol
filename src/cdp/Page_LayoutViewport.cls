VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageLayoutViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Page.LayoutViewport
' Layout viewport position and dimensions.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' pageX : integer
' Horizontal offset relative to the document (CSS pixels).
Public pageX AS integer

' pageY : integer
' Vertical offset relative to the document (CSS pixels).
Public pageY AS integer

' clientWidth : integer
' Width (CSS pixels), excludes scrollbar if present.
Public clientWidth AS integer

' clientHeight : integer
' Height (CSS pixels), excludes scrollbar if present.
Public clientHeight AS integer

Public Function init(ByVal obj as Dictionary) As cdpPageLayoutViewport

    Let pageX = CInt(obj.Item("pageX"))
    Let pageY = CInt(obj.Item("pageY"))
    Let clientWidth = CInt(obj.Item("clientWidth"))
    Let clientHeight = CInt(obj.Item("clientHeight"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("pageX") = pageX
    dict("pageY") = pageY
    dict("clientWidth") = clientWidth
    dict("clientHeight") = clientHeight

    Set asDictionary = dict
End Function
