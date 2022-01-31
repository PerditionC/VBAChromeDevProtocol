VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotTextBoxSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "TextBoxSnapshot"
' DOMSnapshot.TextBoxSnapshot
' Table of details of the post layout rendered text positions. The exact layout should not be regarded as
'   stable and may change between versions.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' layoutIndex : array of integer
' Index of the layout tree node that owns this box collection.
Public layoutIndex AS Collection

' bounds : array of cdpDOMSnapshotRectangle
' The absolute position bounding box.
Public bounds AS Collection

' start : array of integer
' The starting index in characters, for this post layout textbox substring. Characters that
    '   would be represented as a surrogate pair in UTF-16 have length 2.
Public start AS Collection

' length : array of integer
' The number of characters in this post layout textbox substring. Characters that would be
    '   represented as a surrogate pair in UTF-16 have length 2.
Public length AS Collection

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotTextBoxSnapshot
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("layoutIndex")
        layoutIndex.Add v
    Next v
    For each v in obj.Item("bounds")
        bounds.Add v
    Next v
    For each v in obj.Item("start")
        start.Add v
    Next v
    For each v in obj.Item("length")
        length.Add v
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("layoutIndex") = layoutIndex
    Set dict("bounds") = bounds
    Set dict("start") = start
    Set dict("length") = length

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set layoutIndex = New Collection
    Set bounds = New Collection
    Set start = New Collection
    Set length = New Collection
End Sub

