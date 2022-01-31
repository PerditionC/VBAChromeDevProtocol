VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebuggerScriptPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Debugger.ScriptPosition
' Location in the source code.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' lineNumber : integer
' 
Public lineNumber AS integer

' columnNumber : integer
' 
Public columnNumber AS integer

Public Function init(ByVal obj as Dictionary) As cdpDebuggerScriptPosition

    Let lineNumber = CInt(obj.Item("lineNumber"))
    Let columnNumber = CInt(obj.Item("columnNumber"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("lineNumber") = lineNumber
    dict("columnNumber") = columnNumber

    Set asDictionary = dict
End Function

