VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSSourceRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' CSS.SourceRange
' Text range within a resource. All numbers are zero-based.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' startLine : integer
' Start line of range.
Public startLine AS integer

' startColumn : integer
' Start column of range (inclusive).
Public startColumn AS integer

' endLine : integer
' End line of range
Public endLine AS integer

' endColumn : integer
' End column of range (exclusive).
Public endColumn AS integer

Public Function init(ByVal obj as Dictionary) As cdpCSSSourceRange

    Let startLine = CInt(obj.Item("startLine"))
    Let startColumn = CInt(obj.Item("startColumn"))
    Let endLine = CInt(obj.Item("endLine"))
    Let endColumn = CInt(obj.Item("endColumn"))

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
