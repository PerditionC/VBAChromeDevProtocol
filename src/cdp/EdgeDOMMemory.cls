VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpEdgeDOMMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpEdgeDOMMemory"
' EdgeDOMMemory [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' None


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

' Returns list of detached node ids
' Returns: detachedNodesIds - The list of ids
Public Function getDetachedNodesIds() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("EdgeDOMMemory.getDetachedNodesIds", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("detachedNodesIds") Then Set getDetachedNodesIds = results("detachedNodesIds")
    End If
End Function

' Unlinks nodes within the detached node trees to isolate true garbage
' Returns: detachedNodesIds - The list of ids after the tree is unlinked
Public Function unlinkDetachedNodes() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("EdgeDOMMemory.unlinkDetachedNodes", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("detachedNodesIds") Then Set unlinkDetachedNodes = results("detachedNodesIds")
    End If
End Function

' Disables EdgeDOMMemory domain
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "EdgeDOMMemory.disable", params
End Sub

' Enables EdgeDOMMemory domain
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "EdgeDOMMemory.enable", params
End Sub

