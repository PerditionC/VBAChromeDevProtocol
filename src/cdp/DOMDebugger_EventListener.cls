VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMDebuggerEventListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' DOMDebugger.EventListener
' Object event listener.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : string
' `EventListener`'s type.
Public fieldType AS string

' useCapture : boolean
' `EventListener`'s useCapture.
Public useCapture AS boolean

' passive : boolean
' `EventListener`'s passive flag.
Public passive AS boolean

' once : boolean
' `EventListener`'s once flag.
Public once AS boolean

' scriptId : Runtime.ScriptId
' Script id of the handler code.
Public scriptId AS string

' lineNumber : integer
' Line number in the script (0-based).
Public lineNumber AS integer

' columnNumber : integer
' Column number in the script (0-based).
Public columnNumber AS integer

' handler : Runtime.RemoteObject
' Optional
' Event handler function value.
Public handler AS cdpRuntimeRemoteObject

' originalHandler : Runtime.RemoteObject
' Optional
' Event original handler function value.
Public originalHandler AS cdpRuntimeRemoteObject

' backendNodeId : DOM.BackendNodeId
' Optional
' Node the listener is added to (if any).
Public backendNodeId AS integer

Public Function init(ByVal obj as Dictionary) As cdpDOMDebuggerEventListener

    Let fieldType = CStr(obj.Item("type"))
    Let useCapture = CBool(obj.Item("useCapture"))
    Let passive = CBool(obj.Item("passive"))
    Let once = CBool(obj.Item("once"))
    Let scriptId = CStr(obj.Item("scriptId"))
    Let lineNumber = CInt(obj.Item("lineNumber"))
    Let columnNumber = CInt(obj.Item("columnNumber"))
    If obj.Exists("handler") Then
        Set handler = New cdpRuntimeRemoteObject
        handler.init obj.Item("handler")
    End If
    If obj.Exists("originalHandler") Then
        Set originalHandler = New cdpRuntimeRemoteObject
        originalHandler.init obj.Item("originalHandler")
    End If
    If obj.Exists("backendNodeId") Then Let backendNodeId = CInt(obj.Item("backendNodeId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    dict("useCapture") = useCapture
    dict("passive") = passive
    dict("once") = once
    dict("scriptId") = scriptId
    dict("lineNumber") = lineNumber
    dict("columnNumber") = columnNumber
    If Not handler Is Nothing Then Set dict("handler") = handler.asDictionary()
    If Not originalHandler Is Nothing Then Set dict("originalHandler") = originalHandler.asDictionary()
    dict("backendNodeId") = backendNodeId

    Set asDictionary = dict
End Function

