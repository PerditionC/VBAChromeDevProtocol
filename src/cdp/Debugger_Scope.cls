VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebuggerScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Scope"
' Debugger.Scope
' Scope description.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : string
' Scope type.
Public fieldType AS string

' object : Runtime.RemoteObject
' Object representing the scope. For `global` and `with` scopes it represents the actual
    '   object; for the rest of the scopes, it is artificial transient object enumerating scope
    '   variables as its properties.
Public object AS cdpRuntimeRemoteObject

' name : string
' Optional
' 
Public name AS string

' startLocation : Location
' Optional
' Location in the source code where scope starts
Public startLocation AS cdpDebuggerLocation

' endLocation : Location
' Optional
' Location in the source code where scope ends
Public endLocation AS cdpDebuggerLocation

Public Function init(ByVal obj as Dictionary) As cdpDebuggerScope
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let fieldType = CStr(obj.Item("type"))
    Set object = New cdpRuntimeRemoteObject
    object.init obj.Item("object")
    If obj.Exists("name") Then Let name = CStr(obj.Item("name"))
    If obj.Exists("startLocation") Then
        Set startLocation = New cdpDebuggerLocation
        startLocation.init obj.Item("startLocation")
    End If
    If obj.Exists("endLocation") Then
        Set endLocation = New cdpDebuggerLocation
        endLocation.init obj.Item("endLocation")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    Set dict("object") = object.asDictionary()
    if name <> vbNullString Then dict("name") = name
    If Not startLocation Is Nothing Then Set dict("startLocation") = startLocation.asDictionary()
    If Not endLocation Is Nothing Then Set dict("endLocation") = endLocation.asDictionary()

    Set asDictionary = dict
End Function

