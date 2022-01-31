VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeExecutionContextDescr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Runtime.ExecutionContextDescription
' Description of an isolated world.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' id : ExecutionContextId
' Unique id of the execution context. It can be used to specify in which execution context
    '   script evaluation should be performed.
Public id AS integer

' origin : string
' Execution context origin.
Public origin AS string

' name : string
' Human readable name describing given context.
Public name AS string

' uniqueId : string
' A system-unique execution context identifier. Unlike the id, this is unique across
    '   multiple processes, so can be reliably used to identify specific context while backend
    '   performs a cross-process navigation.
Public uniqueId AS string

' auxData : object
' Optional
' Embedder-specific auxiliary data.
Public auxData AS Dictionary

Public Function init(ByVal obj as Dictionary) As cdpRuntimeExecutionContextDescr

    Let id = CInt(obj.Item("id"))
    Let origin = CStr(obj.Item("origin"))
    Let name = CStr(obj.Item("name"))
    Let uniqueId = CStr(obj.Item("uniqueId"))
    If obj.Exists("auxData") Then
        Set auxData = obj.Item("auxData")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("id") = id
    dict("origin") = origin
    dict("name") = name
    dict("uniqueId") = uniqueId
    If Not auxData Is Nothing Then Set dict("auxData") = auxData

    Set asDictionary = dict
End Function
