VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Profile"
' Profiler.Profile
' Profile.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' nodes : array of cdpProfilerProfileNode
' The list of profile nodes. First item is the root node.
Public nodes AS Collection

' startTime : number
' Profiling start timestamp in microseconds.
Public startTime AS Double

' endTime : number
' Profiling end timestamp in microseconds.
Public endTime AS Double

' samples : array of integer
' Optional
' Ids of samples top nodes.
Public samples AS Collection

' timeDeltas : array of integer
' Optional
' Time intervals between adjacent samples in microseconds. The first delta is relative to the
    '   profile startTime.
Public timeDeltas AS Collection

Public Function init(ByVal obj as Dictionary) As cdpProfilerProfile
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("nodes")
        Dim obj_nodes As cdpProfilerProfileNode: Set obj_nodes = New cdpProfilerProfileNode
        obj_nodes.init v
        nodes.Add obj_nodes
        Set obj_nodes = Nothing
    Next v
    Let startTime = CDbl(obj.Item("startTime"))
    Let endTime = CDbl(obj.Item("endTime"))
    If obj.Exists("samples") Then
        For each v in obj.Item("samples")
            samples.Add v
        Next v
    End If
    If obj.Exists("timeDeltas") Then
        For each v in obj.Item("timeDeltas")
            timeDeltas.Add v
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("nodes") = nodes
    dict("startTime") = startTime
    dict("endTime") = endTime
    Set dict("samples") = samples
    Set dict("timeDeltas") = timeDeltas

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set nodes = New Collection
    Set samples = New Collection
    Set timeDeltas = New Collection
End Sub

