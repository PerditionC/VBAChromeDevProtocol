VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerCoverageRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CoverageRange"
' Profiler.CoverageRange
' Coverage data for a source range.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' startOffset : integer
' JavaScript script source offset for the range start.
Public startOffset AS Long

' endOffset : integer
' JavaScript script source offset for the range end.
Public endOffset AS Long

' count : integer
' Collected execution count of the source range.
Public count AS Long

Public Function init(ByVal obj as Dictionary) As cdpProfilerCoverageRange
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let startOffset = CLng(obj.Item("startOffset"))
    Let endOffset = CLng(obj.Item("endOffset"))
    Let count = CLng(obj.Item("count"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("startOffset") = startOffset
    dict("endOffset") = endOffset
    dict("count") = count

    Set asDictionary = dict
End Function

