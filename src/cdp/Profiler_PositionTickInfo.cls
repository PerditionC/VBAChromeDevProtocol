VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerPositionTickInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Profiler.PositionTickInfo
' Specifies a number of samples attributed to a certain source position.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' line : integer
' Source line number (1-based).
Public line AS integer

' ticks : integer
' Number of samples attributed to the source line.
Public ticks AS integer

Public Function init(ByVal obj as Dictionary) As cdpProfilerPositionTickInfo

    Let line = CInt(obj.Item("line"))
    Let ticks = CInt(obj.Item("ticks"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("line") = line
    dict("ticks") = ticks

    Set asDictionary = dict
End Function
