VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebuggerDebugSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DebugSymbols"
' Debugger.DebugSymbols
' Debug symbols available for a wasm script.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : string
' Type of the debug symbols.
Public fieldType AS string

' externalURL : string
' Optional
' URL of the external symbol source.
Public externalURL AS string

Public Function init(ByVal obj as Dictionary) As cdpDebuggerDebugSymbols
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let fieldType = CStr(obj.Item("type"))
    If obj.Exists("externalURL") Then Let externalURL = CStr(obj.Item("externalURL"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    if externalURL <> vbNullString Then dict("externalURL") = externalURL

    Set asDictionary = dict
End Function

