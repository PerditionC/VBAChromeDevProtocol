VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeCallFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CallFrame"
' Runtime.CallFrame
' Stack entry for runtime errors and assertions.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' functionName : string
' JavaScript function name.
Public functionName AS string

' scriptId : ScriptId
' JavaScript script id.
Public scriptId AS string

' url : string
' JavaScript script name or url.
Public url AS string

' lineNumber : integer
' JavaScript script line number (0-based).
Public lineNumber AS Long

' columnNumber : integer
' JavaScript script column number (0-based).
Public columnNumber AS Long

Public Function init(ByVal obj as Dictionary) As cdpRuntimeCallFrame
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let functionName = CStr(obj.Item("functionName"))
    Let scriptId = CStr(obj.Item("scriptId"))
    Let url = CStr(obj.Item("url"))
    Let lineNumber = CLng(obj.Item("lineNumber"))
    Let columnNumber = CLng(obj.Item("columnNumber"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("functionName") = functionName
    dict("scriptId") = scriptId
    dict("url") = url
    dict("lineNumber") = lineNumber
    dict("columnNumber") = columnNumber

    Set asDictionary = dict
End Function

