VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMRGBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "RGBA"
' DOM.RGBA
' A structure holding an RGBA color.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' r : integer
' The red component, in the [0-255] range.
Public r AS Long

' g : integer
' The green component, in the [0-255] range.
Public g AS Long

' b : integer
' The blue component, in the [0-255] range.
Public b AS Long

' a : number
' Optional
' The alpha component, in the [0-1] range (default: 1).
Public a AS Double

Public Function init(ByVal obj as Dictionary) As cdpDOMRGBA
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let r = CLng(obj.Item("r"))
    Let g = CLng(obj.Item("g"))
    Let b = CLng(obj.Item("b"))
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

