VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebuggerSearchMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SearchMatch"
' Debugger.SearchMatch
' Search match for resource.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' lineNumber : number
' Line number in resource content.
Public lineNumber AS Double

' lineContent : string
' Line with match content.
Public lineContent AS string

Public Function init(ByVal obj as Dictionary) As cdpDebuggerSearchMatch
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let lineNumber = CDbl(obj.Item("lineNumber"))
    Let lineContent = CStr(obj.Item("lineContent"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("lineNumber") = lineNumber
    dict("lineContent") = lineContent

    Set asDictionary = dict
End Function

