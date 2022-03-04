VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebuggerLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Location"
' Debugger.Location
' Location in the source code.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' scriptId : Runtime.ScriptId
' Script identifier as reported in the `Debugger.scriptParsed`.
Public scriptId AS string

' lineNumber : integer
' Line number in the script (0-based).
Public lineNumber AS Long

' columnNumber : integer
' Optional
' Column number in the script (0-based).
Public columnNumber AS Long

Public Function init(ByVal obj as Dictionary) As cdpDebuggerLocation
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let scriptId = CStr(obj.Item("scriptId"))
    Let lineNumber = CLng(obj.Item("lineNumber"))
    If obj.Exists("columnNumber") Then Let columnNumber = CLng(obj.Item("columnNumber"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("scriptId") = scriptId
    dict("lineNumber") = lineNumber
    dict("columnNumber") = columnNumber

    Set asDictionary = dict
End Function

