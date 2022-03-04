VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMBoxModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "BoxModel"
' DOM.BoxModel
' Box model.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' content : Quad
' Content box
Public content AS Collection

' padding : Quad
' Padding box
Public padding AS Collection

' border : Quad
' Border box
Public border AS Collection

' margin : Quad
' Margin box
Public margin AS Collection

' width : integer
' Node width
Public width AS Long

' height : integer
' Node height
Public height AS Long

' shapeOutside : ShapeOutsideInfo
' Optional
' Shape outside coordinates
Public shapeOutside AS cdpDOMShapeOutsideInfo

Public Function init(ByVal obj as Dictionary) As cdpDOMBoxModel
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("content")
        content.Add v
    Next v
    For each v in obj.Item("padding")
        padding.Add v
    Next v
    For each v in obj.Item("border")
        border.Add v
    Next v
    For each v in obj.Item("margin")
        margin.Add v
    Next v
    Let width = CLng(obj.Item("width"))
    Let height = CLng(obj.Item("height"))
    If obj.Exists("shapeOutside") Then
        Set shapeOutside = New cdpDOMShapeOutsideInfo
        shapeOutside.init obj.Item("shapeOutside")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("content") = content
    Set dict("padding") = padding
    Set dict("border") = border
    Set dict("margin") = margin
    dict("width") = width
    dict("height") = height
    If Not shapeOutside Is Nothing Then Set dict("shapeOutside") = shapeOutside.asDictionary()

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set content = New Collection
    Set padding = New Collection
    Set border = New Collection
    Set margin = New Collection
End Sub

