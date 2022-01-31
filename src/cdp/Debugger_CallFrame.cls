VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebuggerCallFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CallFrame"
' Debugger.CallFrame
' JavaScript call frame. Array of call frames form the call stack.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' callFrameId : CallFrameId
' Call frame identifier. This identifier is only valid while the virtual machine is paused.
Public callFrameId AS string

' functionName : string
' Name of the JavaScript function called on this call frame.
Public functionName AS string

' functionLocation : Location
' Optional
' Location in the source code.
Public functionLocation AS cdpDebuggerLocation

' location : Location
' Location in the source code.
Public location AS cdpDebuggerLocation

' url : string
' JavaScript script name or url.
Public url AS string

' scopeChain : array of cdpDebuggerScope
' Scope chain for this call frame.
Public scopeChain AS Collection

' this : Runtime.RemoteObject
' `this` object for this call frame.
Public this AS cdpRuntimeRemoteObject

' returnValue : Runtime.RemoteObject
' Optional
' The value being returned, if the function is at return point.
Public returnValue AS cdpRuntimeRemoteObject

Public Function init(ByVal obj as Dictionary) As cdpDebuggerCallFrame
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let callFrameId = CStr(obj.Item("callFrameId"))
    Let functionName = CStr(obj.Item("functionName"))
    If obj.Exists("functionLocation") Then
        Set functionLocation = New cdpDebuggerLocation
        functionLocation.init obj.Item("functionLocation")
    End If
    Set location = New cdpDebuggerLocation
    location.init obj.Item("location")
    Let url = CStr(obj.Item("url"))
    For each v in obj.Item("scopeChain")
        Dim obj_scopeChain As cdpDebuggerScope: Set obj_scopeChain = New cdpDebuggerScope
        obj_scopeChain.init v
        scopeChain.Add obj_scopeChain
        Set obj_scopeChain = Nothing
    Next v
    Set this = New cdpRuntimeRemoteObject
    this.init obj.Item("this")
    If obj.Exists("returnValue") Then
        Set returnValue = New cdpRuntimeRemoteObject
        returnValue.init obj.Item("returnValue")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("callFrameId") = callFrameId
    dict("functionName") = functionName
    If Not functionLocation Is Nothing Then Set dict("functionLocation") = functionLocation.asDictionary()
    Set dict("location") = location.asDictionary()
    dict("url") = url
    Set dict("scopeChain") = scopeChain
    Set dict("this") = this.asDictionary()
    If Not returnValue Is Nothing Then Set dict("returnValue") = returnValue.asDictionary()

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set scopeChain = New Collection
End Sub

