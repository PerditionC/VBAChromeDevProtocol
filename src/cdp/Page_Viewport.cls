VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Viewport"
' Page.Viewport
' Viewport for capturing screenshot.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' x : number
' X offset in device independent pixels (dip).
Public x AS Double

' y : number
' Y offset in device independent pixels (dip).
Public y AS Double

' width : number
' Rectangle width in device independent pixels (dip).
Public width AS Double

' height : number
' Rectangle height in device independent pixels (dip).
Public height AS Double

' scale : number
' Page scale factor.
Public fieldScale AS Double

Public Function init(ByVal obj as Dictionary) As cdpPageViewport
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let x = CDbl(obj.Item("x"))
    Let y = CDbl(obj.Item("y"))
    Let width = CDbl(obj.Item("width"))
    Let height = CDbl(obj.Item("height"))
    Let fieldScale = CDbl(obj.Item("scale"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("x") = x
    dict("y") = y
    dict("width") = width
    dict("height") = height
    dict("scale") = fieldScale

    Set asDictionary = dict
End Function

