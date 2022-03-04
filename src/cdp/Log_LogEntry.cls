VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpLogLogEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "LogEntry"
' Log.LogEntry
' Log entry.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' source : string
' Log entry source.
Public source AS string

' level : string
' Log entry severity.
Public level AS string

' text : string
' Logged text.
Public text AS string

' category : string
' Optional
' 
Public category AS string

' timestamp : Runtime.Timestamp
' Timestamp when this entry was added.
Public timestamp AS Double

' url : string
' Optional
' URL of the resource if known.
Public url AS string

' lineNumber : integer
' Optional
' Line number in the resource.
Public lineNumber AS Long

' stackTrace : Runtime.StackTrace
' Optional
' JavaScript stack trace.
Public stackTrace AS cdpRuntimeStackTrace

' networkRequestId : Network.RequestId
' Optional
' Identifier of the network request associated with this entry.
Public networkRequestId AS string

' workerId : string
' Optional
' Identifier of the worker associated with this entry.
Public workerId AS string

' args : array of cdpRuntimeRemoteObject
' Optional
' Call arguments.
Public args AS Collection

Public Function init(ByVal obj as Dictionary) As cdpLogLogEntry
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let source = CStr(obj.Item("source"))
    Let level = CStr(obj.Item("level"))
    Let text = CStr(obj.Item("text"))
    If obj.Exists("category") Then Let category = CStr(obj.Item("category"))
    Let timestamp = CDbl(obj.Item("timestamp"))
    If obj.Exists("url") Then Let url = CStr(obj.Item("url"))
    If obj.Exists("lineNumber") Then Let lineNumber = CLng(obj.Item("lineNumber"))
    If obj.Exists("stackTrace") Then
        Set stackTrace = New cdpRuntimeStackTrace
        stackTrace.init obj.Item("stackTrace")
    End If
    If obj.Exists("networkRequestId") Then Let networkRequestId = CStr(obj.Item("networkRequestId"))
    If obj.Exists("workerId") Then Let workerId = CStr(obj.Item("workerId"))
    If obj.Exists("args") Then
        For each v in obj.Item("args")
            Dim obj_args As cdpRuntimeRemoteObject: Set obj_args = New cdpRuntimeRemoteObject
            obj_args.init v
            args.Add obj_args
            Set obj_args = Nothing
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("source") = source
    dict("level") = level
    dict("text") = text
    if category <> vbNullString Then dict("category") = category
    dict("timestamp") = timestamp
    if url <> vbNullString Then dict("url") = url
    dict("lineNumber") = lineNumber
    If Not stackTrace Is Nothing Then Set dict("stackTrace") = stackTrace.asDictionary()
    if networkRequestId <> vbNullString Then dict("networkRequestId") = networkRequestId
    if workerId <> vbNullString Then dict("workerId") = workerId
    Set dict("args") = args

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set args = New Collection
End Sub

