VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeStackTraceId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "StackTraceId"
' Runtime.StackTraceId
' If `debuggerId` is set stack trace comes from another debugger and can be resolved there. This
'   allows to track cross-debugger calls. See `Runtime.StackTrace` and `Debugger.paused` for usages.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' id : string
' 
Public id AS string

' debuggerId : UniqueDebuggerId
' Optional
' 
Public debuggerId AS string

Public Function init(ByVal obj as Dictionary) As cdpRuntimeStackTraceId
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let id = CStr(obj.Item("id"))
    If obj.Exists("debuggerId") Then Let debuggerId = CStr(obj.Item("debuggerId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("id") = id
    if debuggerId <> vbNullString Then dict("debuggerId") = debuggerId

    Set asDictionary = dict
End Function

