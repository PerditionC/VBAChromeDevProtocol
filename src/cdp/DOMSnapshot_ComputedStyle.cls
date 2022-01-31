VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotComputedStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ComputedStyle"
' DOMSnapshot.ComputedStyle
' A subset of the full ComputedStyle as defined by the request whitelist.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' properties : array of cdpDOMSnapshotNameValue
' Name/value pairs of computed style properties.
Public properties AS Collection

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotComputedStyle
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("properties")
        Dim obj_properties As cdpDOMSnapshotNameValue: Set obj_properties = New cdpDOMSnapshotNameValue
        obj_properties.init v
        properties.Add obj_properties
        Set obj_properties = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("properties") = properties

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set properties = New Collection
End Sub

