VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeExceptionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ExceptionDetails"
' Runtime.ExceptionDetails
' Detailed information about exception (or error) that was thrown during script compilation or
'   execution.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' exceptionId : integer
' Exception id.
Public exceptionId AS Long

' text : string
' Exception text, which should be used together with exception object when available.
Public text AS string

' lineNumber : integer
' Line number of the exception location (0-based).
Public lineNumber AS Long

' columnNumber : integer
' Column number of the exception location (0-based).
Public columnNumber AS Long

' scriptId : ScriptId
' Optional
' Script ID of the exception location.
Public scriptId AS string

' url : string
' Optional
' URL of the exception location, to be used when the script was not reported.
Public url AS string

' stackTrace : StackTrace
' Optional
' JavaScript stack trace if available.
Public stackTrace AS cdpRuntimeStackTrace

' exception : RemoteObject
' Optional
' Exception object if available.
Public exception AS cdpRuntimeRemoteObject

' executionContextId : ExecutionContextId
' Optional
' Identifier of the context where exception happened.
Public executionContextId AS Long

' exceptionMetaData : object
' Optional
' Dictionary with entries of meta data that the client associated
    '   with this exception, such as information about associated network
    '   requests, etc.
Public exceptionMetaData AS Dictionary

Public Function init(ByVal obj as Dictionary) As cdpRuntimeExceptionDetails
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let exceptionId = CLng(obj.Item("exceptionId"))
    Let text = CStr(obj.Item("text"))
    Let lineNumber = CLng(obj.Item("lineNumber"))
    Let columnNumber = CLng(obj.Item("columnNumber"))
    If obj.Exists("scriptId") Then Let scriptId = CStr(obj.Item("scriptId"))
    If obj.Exists("url") Then Let url = CStr(obj.Item("url"))
    If obj.Exists("stackTrace") Then
        Set stackTrace = New cdpRuntimeStackTrace
        stackTrace.init obj.Item("stackTrace")
    End If
    If obj.Exists("exception") Then
        Set exception = New cdpRuntimeRemoteObject
        exception.init obj.Item("exception")
    End If
    If obj.Exists("executionContextId") Then Let executionContextId = CLng(obj.Item("executionContextId"))
    If obj.Exists("exceptionMetaData") Then
        Set exceptionMetaData = obj.Item("exceptionMetaData")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("exceptionId") = exceptionId
    dict("text") = text
    dict("lineNumber") = lineNumber
    dict("columnNumber") = columnNumber
    if scriptId <> vbNullString Then dict("scriptId") = scriptId
    if url <> vbNullString Then dict("url") = url
    If Not stackTrace Is Nothing Then Set dict("stackTrace") = stackTrace.asDictionary()
    If Not exception Is Nothing Then Set dict("exception") = exception.asDictionary()
    dict("executionContextId") = executionContextId
    If Not exceptionMetaData Is Nothing Then Set dict("exceptionMetaData") = exceptionMetaData

    Set asDictionary = dict
End Function

