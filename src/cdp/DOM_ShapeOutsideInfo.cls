VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMShapeOutsideInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ShapeOutsideInfo"
' DOM.ShapeOutsideInfo
' CSS Shape Outside details.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' bounds : Quad
' Shape bounds
Public bounds AS Collection

' shape : array of any
' Shape coordinate details
Public shape AS Collection

' marginShape : array of any
' Margin shape bounds
Public marginShape AS Collection

Public Function init(ByVal obj as Dictionary) As cdpDOMShapeOutsideInfo
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("bounds")
        bounds.Add v
    Next v
    For each v in obj.Item("shape")
        shape.Add v
    Next v
    For each v in obj.Item("marginShape")
        marginShape.Add v
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("bounds") = bounds
    Set dict("shape") = shape
    Set dict("marginShape") = marginShape

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set bounds = New Collection
    Set shape = New Collection
    Set marginShape = New Collection
End Sub

