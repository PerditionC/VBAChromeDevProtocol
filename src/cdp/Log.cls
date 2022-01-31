VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpLog"
' Log
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Log entry.
' LogEntry AS object

' Violation configuration setting.
' ViolationSetting AS object



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

' Clears the log.
Public Sub clear()
    Dim params As New Dictionary
    cdp.InvokeMethod "Log.clear", params
End Sub

' Disables log domain, prevents further log entries from being reported to the client.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Log.disable", params
End Sub

' Enables log domain, sends the entries collected so far to the client by means of the
' `entryAdded` notification.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Log.enable", params
End Sub

' start violation reporting.
Public Sub startViolationsReport( _
    ByVal config AS Collection _
)
    ' config: array Configuration for violations.

    Dim params As New Dictionary
    Set params("config") = config

    cdp.InvokeMethod "Log.startViolationsReport", params
End Sub

' Stop violation reporting.
Public Sub stopViolationsReport()
    Dim params As New Dictionary
    cdp.InvokeMethod "Log.stopViolationsReport", params
End Sub

