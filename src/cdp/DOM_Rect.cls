VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Rect"
' DOM.Rect
' Rectangle.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' x : number
' X coordinate
Public x AS Double

' y : number
' Y coordinate
Public y AS Double

' width : number
' Rectangle width
Public width AS Double

' height : number
' Rectangle height
Public height AS Double

Public Function init(ByVal obj as Dictionary) As cdpDOMRect
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let x = CDbl(obj.Item("x"))
    Let y = CDbl(obj.Item("y"))
    Let width = CDbl(obj.Item("width"))
    Let height = CDbl(obj.Item("height"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("x") = x
    dict("y") = y
    dict("width") = width
    dict("height") = height

    Set asDictionary = dict
End Function

