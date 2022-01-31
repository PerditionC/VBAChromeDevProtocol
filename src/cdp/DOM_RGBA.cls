VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMRGBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' DOM.RGBA
' A structure holding an RGBA color.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' r : integer
' The red component, in the [0-255] range.
Public r AS integer

' g : integer
' The green component, in the [0-255] range.
Public g AS integer

' b : integer
' The blue component, in the [0-255] range.
Public b AS integer

' a : number
' Optional
' The alpha component, in the [0-1] range (default: 1).
Public a AS Double

Public Function init(ByVal obj as Dictionary) As cdpDOMRGBA

    Let r = CInt(obj.Item("r"))
    Let g = CInt(obj.Item("g"))
    Let b = CInt(obj.Item("b"))
    If obj.Exists("a") Then Let a = CDbl(obj.Item("a"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("r") = r
    dict("g") = g
    dict("b") = b
    dict("a") = a

    Set asDictionary = dict
End Function

