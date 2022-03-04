VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageLayoutViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "LayoutViewport"
' Page.LayoutViewport
' Layout viewport position and dimensions.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' pageX : integer
' Horizontal offset relative to the document (CSS pixels).
Public pageX AS Long

' pageY : integer
' Vertical offset relative to the document (CSS pixels).
Public pageY AS Long

' clientWidth : integer
' Width (CSS pixels), excludes scrollbar if present.
Public clientWidth AS Long

' clientHeight : integer
' Height (CSS pixels), excludes scrollbar if present.
Public clientHeight AS Long

Public Function init(ByVal obj as Dictionary) As cdpPageLayoutViewport
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let pageX = CLng(obj.Item("pageX"))
    Let pageY = CLng(obj.Item("pageY"))
    Let clientWidth = CLng(obj.Item("clientWidth"))
    Let clientHeight = CLng(obj.Item("clientHeight"))

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

