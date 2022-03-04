VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSSourceRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SourceRange"
' CSS.SourceRange
' Text range within a resource. All numbers are zero-based.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' startLine : integer
' Start line of range.
Public startLine AS Long

' startColumn : integer
' Start column of range (inclusive).
Public startColumn AS Long

' endLine : integer
' End line of range
Public endLine AS Long

' endColumn : integer
' End column of range (exclusive).
Public endColumn AS Long

Public Function init(ByVal obj as Dictionary) As cdpCSSSourceRange
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let startLine = CLng(obj.Item("startLine"))
    Let startColumn = CLng(obj.Item("startColumn"))
    Let endLine = CLng(obj.Item("endLine"))
    Let endColumn = CLng(obj.Item("endColumn"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("startLine") = startLine
    dict("startColumn") = startColumn
    dict("endLine") = endLine
    dict("endColumn") = endColumn

    Set asDictionary = dict
End Function

