VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebuggerLocationRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "LocationRange"
' Debugger.LocationRange
' Location range within one script.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' scriptId : Runtime.ScriptId
' 
Public scriptId AS string

' start : ScriptPosition
' 
Public start AS cdpDebuggerScriptPosition

' end : ScriptPosition
' 
Public fieldEnd AS cdpDebuggerScriptPosition

Public Function init(ByVal obj as Dictionary) As cdpDebuggerLocationRange
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let scriptId = CStr(obj.Item("scriptId"))
    Set start = New cdpDebuggerScriptPosition
    start.init obj.Item("start")
    Set fieldEnd = New cdpDebuggerScriptPosition
    fieldEnd.init obj.Item("end")

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("scriptId") = scriptId
    Set dict("start") = start.asDictionary()
    Set dict("end") = fieldEnd.asDictionary()

    Set asDictionary = dict
End Function

