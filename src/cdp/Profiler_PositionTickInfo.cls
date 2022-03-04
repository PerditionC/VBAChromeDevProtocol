VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerPositionTickInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PositionTickInfo"
' Profiler.PositionTickInfo
' Specifies a number of samples attributed to a certain source position.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' line : integer
' Source line number (1-based).
Public line AS Long

' ticks : integer
' Number of samples attributed to the source line.
Public ticks AS Long

Public Function init(ByVal obj as Dictionary) As cdpProfilerPositionTickInfo
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let line = CLng(obj.Item("line"))
    Let ticks = CLng(obj.Item("ticks"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("line") = line
    dict("ticks") = ticks

    Set asDictionary = dict
End Function

