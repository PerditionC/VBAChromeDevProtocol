VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotInlineTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "InlineTextBox"
' DOMSnapshot.InlineTextBox
' Details of post layout rendered text positions. The exact layout should not be regarded as
'   stable and may change between versions.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' boundingBox : DOM.Rect
' The bounding box in document coordinates. Note that scroll offset of the document is ignored.
Public boundingBox AS cdpDOMRect

' startCharacterIndex : integer
' The starting index in characters, for this post layout textbox substring. Characters that
    '   would be represented as a surrogate pair in UTF-16 have length 2.
Public startCharacterIndex AS Long

' numCharacters : integer
' The number of characters in this post layout textbox substring. Characters that would be
    '   represented as a surrogate pair in UTF-16 have length 2.
Public numCharacters AS Long

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotInlineTextBox
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Set boundingBox = New cdpDOMRect
    boundingBox.init obj.Item("boundingBox")
    Let startCharacterIndex = CLng(obj.Item("startCharacterIndex"))
    Let numCharacters = CLng(obj.Item("numCharacters"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("boundingBox") = boundingBox.asDictionary()
    dict("startCharacterIndex") = startCharacterIndex
    dict("numCharacters") = numCharacters

    Set asDictionary = dict
End Function

