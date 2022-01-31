VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpMemorySamplingProfileNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SamplingProfileNode"
' Memory.SamplingProfileNode
' Heap profile sample.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' size : number
' Size of the sampled allocation.
Public size AS Double

' total : number
' Total bytes attributed to this sample.
Public total AS Double

' stack : array of string
' Execution stack at the point of allocation.
Public stack AS Collection

Public Function init(ByVal obj as Dictionary) As cdpMemorySamplingProfileNode
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let size = CDbl(obj.Item("size"))
    Let total = CDbl(obj.Item("total"))
    For each v in obj.Item("stack")
        stack.Add v
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("size") = size
    dict("total") = total
    Set dict("stack") = stack

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set stack = New Collection
End Sub

