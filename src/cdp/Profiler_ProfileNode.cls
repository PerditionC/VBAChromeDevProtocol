VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerProfileNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ProfileNode"
' Profiler.ProfileNode
' Profile node. Holds callsite information, execution statistics and child nodes.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' id : integer
' Unique id of the node.
Public id AS Long

' callFrame : Runtime.CallFrame
' Function location.
Public callFrame AS cdpRuntimeCallFrame

' hitCount : integer
' Optional
' Number of samples where this node was on top of the call stack.
Public hitCount AS Long

' children : array of integer
' Optional
' Child node ids.
Public children AS Collection

' deoptReason : string
' Optional
' The reason of being not optimized. The function may be deoptimized or marked as don't
    '   optimize.
Public deoptReason AS string

' positionTicks : array of cdpProfilerPositionTickInfo
' Optional
' An array of source position ticks.
Public positionTicks AS Collection

Public Function init(ByVal obj as Dictionary) As cdpProfilerProfileNode
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let id = CLng(obj.Item("id"))
    Set callFrame = New cdpRuntimeCallFrame
    callFrame.init obj.Item("callFrame")
    If obj.Exists("hitCount") Then Let hitCount = CLng(obj.Item("hitCount"))
    If obj.Exists("children") Then
        For each v in obj.Item("children")
            children.Add v
        Next v
    End If
    If obj.Exists("deoptReason") Then Let deoptReason = CStr(obj.Item("deoptReason"))
    If obj.Exists("positionTicks") Then
        For each v in obj.Item("positionTicks")
            Dim obj_positionTicks As cdpProfilerPositionTickInfo: Set obj_positionTicks = New cdpProfilerPositionTickInfo
            obj_positionTicks.init v
            positionTicks.Add obj_positionTicks
            Set obj_positionTicks = Nothing
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("id") = id
    Set dict("callFrame") = callFrame.asDictionary()
    dict("hitCount") = hitCount
    Set dict("children") = children
    if deoptReason <> vbNullString Then dict("deoptReason") = deoptReason
    Set dict("positionTicks") = positionTicks

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set children = New Collection
    Set positionTicks = New Collection
End Sub

