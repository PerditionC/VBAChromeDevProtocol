VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebuggerScriptPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScriptPosition"
' Debugger.ScriptPosition
' Location in the source code.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' lineNumber : integer
' 
Public lineNumber AS Long

' columnNumber : integer
' 
Public columnNumber AS Long

Public Function init(ByVal obj as Dictionary) As cdpDebuggerScriptPosition
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let lineNumber = CLng(obj.Item("lineNumber"))
    Let columnNumber = CLng(obj.Item("columnNumber"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("lineNumber") = lineNumber
    dict("columnNumber") = columnNumber

    Set asDictionary = dict
End Function

