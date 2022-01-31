VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpLayerTreePictureTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PictureTile"
' LayerTree.PictureTile
' Serialized fragment of layer picture along with its offset within the layer.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' x : number
' Offset from owning layer left boundary
Public x AS Double

' y : number
' Offset from owning layer top boundary
Public y AS Double

' picture : binary
' Base64-encoded snapshot data.
Public picture AS String

Public Function init(ByVal obj as Dictionary) As cdpLayerTreePictureTile
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let x = CDbl(obj.Item("x"))
    Let y = CDbl(obj.Item("y"))
    Let picture = CStr(obj.Item("picture"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("x") = x
    dict("y") = y
    dict("picture") = picture

    Set asDictionary = dict
End Function

