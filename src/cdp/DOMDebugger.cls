VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpDOMDebugger"
' DOMDebugger
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' DOM breakpoint type.
' DOMBreakpointType AS string
Public Enum DOMBreakpointType
    DOMBT_subtreemodified
    DOMBT_attributemodified
    DOMBT_noderemoved
End Enum

' CSP Violation type.
' CSPViolationType AS string
Public Enum CSPViolationType
    CSPVT_trustedtypesinkviolation
    CSPVT_trustedtypepolicyviolation
End Enum

' Object event listener.
' EventListener AS object


Public Function DOMBreakpointTypeToString(ByVal dombt As DOMBreakpointType) As String
    Dim retVal As String
    Select Case dombt
        Case DOMBT_subtreemodified
            retVal = "subtree-modified"
        Case DOMBT_attributemodified
            retVal = "attribute-modified"
        Case DOMBT_noderemoved
            retVal = "node-removed"
        Case Else
            Debug.Print "Warning, unknown value " & dombt
    End Select
    DOMBreakpointTypeToString = retVal
End Function

Public Function StringToDOMBreakpointType(ByVal s As String) As DOMBreakpointType
    Dim retVal As DOMBreakpointType
    Select Case s
        Case "subtree-modified"
            retVal = DOMBT_subtreemodified
        Case "attribute-modified"
            retVal = DOMBT_attributemodified
        Case "node-removed"
            retVal = DOMBT_noderemoved
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToDOMBreakpointType = retVal
End Function


Public Function CSPViolationTypeToString(ByVal cspvt As CSPViolationType) As String
    Dim retVal As String
    Select Case cspvt
        Case CSPVT_trustedtypesinkviolation
            retVal = "trustedtype-sink-violation"
        Case CSPVT_trustedtypepolicyviolation
            retVal = "trustedtype-policy-violation"
        Case Else
            Debug.Print "Warning, unknown value " & cspvt
    End Select
    CSPViolationTypeToString = retVal
End Function

Public Function StringToCSPViolationType(ByVal s As String) As CSPViolationType
    Dim retVal As CSPViolationType
    Select Case s
        Case "trustedtype-sink-violation"
            retVal = CSPVT_trustedtypesinkviolation
        Case "trustedtype-policy-violation"
            retVal = CSPVT_trustedtypepolicyviolation
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCSPViolationType = retVal
End Function


Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' Returns event listeners of the given object.
' Returns: listeners - Array of relevant listeners.
Public Function getEventListeners( _
    ByVal objectId AS string, _
    Optional ByVal depth AS Variant, _
    Optional ByVal pierce AS Variant _
) AS Collection
    ' objectId: Runtime.RemoteObjectId Identifier of the object to return listeners for.
    ' depth: integer(optional) The maximum depth at which Node children should be retrieved, defaults to 1. Use -1 for the
    '   entire subtree or provide an integer larger than 0.
    ' pierce: boolean(optional) Whether or not iframes and shadow roots should be traversed when returning the subtree
    '   (default is false). Reports listeners for all contexts if pierce is enabled.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)
    If Not IsMissing(depth) Then params("depth") = CLng(depth)
    If Not IsMissing(pierce) Then params("pierce") = CBool(pierce)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOMDebugger.getEventListeners", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("listeners") Then Set getEventListeners = results("listeners")
    End If
End Function

' Removes DOM breakpoint that was set using `setDOMBreakpoint`.
Public Sub removeDOMBreakpoint( _
    ByVal nodeId AS Long, _
    ByVal pType AS string _
)
    ' nodeId: DOM.NodeId Identifier of the node to remove breakpoint from.
    ' pType: DOMBreakpointType Type of the breakpoint to remove.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("type") = CStr(pType)

    cdp.InvokeMethod "DOMDebugger.removeDOMBreakpoint", params
End Sub

' Removes breakpoint on particular DOM event.
Public Sub removeEventListenerBreakpoint( _
    ByVal eventName AS string, _
    Optional ByVal targetName AS Variant _
)
    ' eventName: string Event name.
    ' targetName: string(optional) EventTarget interface name.

    Dim params As New Dictionary
    params("eventName") = CStr(eventName)
    If Not IsMissing(targetName) Then params("targetName") = CStr(targetName)

    cdp.InvokeMethod "DOMDebugger.removeEventListenerBreakpoint", params
End Sub

' Removes breakpoint on particular native event.
' Experimental
Public Sub removeInstrumentationBreakpoint( _
    ByVal eventName AS string _
)
    ' eventName: string Instrumentation name to stop on.

    Dim params As New Dictionary
    params("eventName") = CStr(eventName)

    cdp.InvokeMethod "DOMDebugger.removeInstrumentationBreakpoint", params
End Sub

' Removes breakpoint from XMLHttpRequest.
Public Sub removeXHRBreakpoint( _
    ByVal url AS string _
)
    ' url: string Resource URL substring.

    Dim params As New Dictionary
    params("url") = CStr(url)

    cdp.InvokeMethod "DOMDebugger.removeXHRBreakpoint", params
End Sub

' Sets breakpoint on particular CSP violations.
' Experimental
Public Sub setBreakOnCSPViolation( _
    ByVal violationTypes AS Collection _
)
    ' violationTypes: array CSP Violations to stop upon.

    Dim params As New Dictionary
    Set params("violationTypes") = violationTypes

    cdp.InvokeMethod "DOMDebugger.setBreakOnCSPViolation", params
End Sub

' Sets breakpoint on particular operation with DOM.
Public Sub setDOMBreakpoint( _
    ByVal nodeId AS Long, _
    ByVal pType AS string _
)
    ' nodeId: DOM.NodeId Identifier of the node to set breakpoint on.
    ' pType: DOMBreakpointType Type of the operation to stop upon.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("type") = CStr(pType)

    cdp.InvokeMethod "DOMDebugger.setDOMBreakpoint", params
End Sub

' Sets breakpoint on particular DOM event.
Public Sub setEventListenerBreakpoint( _
    ByVal eventName AS string, _
    Optional ByVal targetName AS Variant _
)
    ' eventName: string DOM Event name to stop on (any DOM event will do).
    ' targetName: string(optional) EventTarget interface name to stop on. If equal to `"*"` or not provided, will stop on any
    '   EventTarget.

    Dim params As New Dictionary
    params("eventName") = CStr(eventName)
    If Not IsMissing(targetName) Then params("targetName") = CStr(targetName)

    cdp.InvokeMethod "DOMDebugger.setEventListenerBreakpoint", params
End Sub

' Sets breakpoint on particular native event.
' Experimental
Public Sub setInstrumentationBreakpoint( _
    ByVal eventName AS string _
)
    ' eventName: string Instrumentation name to stop on.

    Dim params As New Dictionary
    params("eventName") = CStr(eventName)

    cdp.InvokeMethod "DOMDebugger.setInstrumentationBreakpoint", params
End Sub

' Sets breakpoint on XMLHttpRequest.
Public Sub setXHRBreakpoint( _
    ByVal url AS string _
)
    ' url: string Resource URL substring. All XHRs having this substring in the URL will get stopped upon.

    Dim params As New Dictionary
    params("url") = CStr(url)

    cdp.InvokeMethod "DOMDebugger.setXHRBreakpoint", params
End Sub

