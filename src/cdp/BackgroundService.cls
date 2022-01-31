VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBackgroundService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpBackgroundService"
' BackgroundService [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' The Background Service that will be associated with the commands/events.
'   Every Background Service operates independently, but they share the same
'   API.
' ServiceName AS string
Public Enum ServiceName
    SN_backgroundFetch
    SN_backgroundSync
    SN_pushMessaging
    SN_notifications
    SN_paymentHandler
    SN_periodicBackgroundSync
End Enum

' A key-value pair for additional event information to pass along.
' EventMetadata AS object

' 
' BackgroundServiceEvent AS object


Public Function ServiceNameToString(ByVal sn As ServiceName) As String
    Dim retVal As String
    Select Case sn
        Case SN_backgroundFetch
            retVal = "backgroundFetch"
        Case SN_backgroundSync
            retVal = "backgroundSync"
        Case SN_pushMessaging
            retVal = "pushMessaging"
        Case SN_notifications
            retVal = "notifications"
        Case SN_paymentHandler
            retVal = "paymentHandler"
        Case SN_periodicBackgroundSync
            retVal = "periodicBackgroundSync"
        Case Else
            Debug.Print "Warning, unknown value " & sn
    End Select
    ServiceNameToString = retVal
End Function

Public Function StringToServiceName(ByVal s As String) As ServiceName
    Dim retVal As ServiceName
    Select Case s
        Case "backgroundFetch"
            retVal = SN_backgroundFetch
        Case "backgroundSync"
            retVal = SN_backgroundSync
        Case "pushMessaging"
            retVal = SN_pushMessaging
        Case "notifications"
            retVal = SN_notifications
        Case "paymentHandler"
            retVal = SN_paymentHandler
        Case "periodicBackgroundSync"
            retVal = SN_periodicBackgroundSync
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToServiceName = retVal
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

' Enables event updates for the service.
Public Sub startObserving( _
    ByVal service AS string _
)
    ' service: ServiceName 

    Dim params As New Dictionary
    params("service") = CStr(service)

    cdp.InvokeMethod "BackgroundService.startObserving", params
End Sub

' Disables event updates for the service.
Public Sub stopObserving( _
    ByVal service AS string _
)
    ' service: ServiceName 

    Dim params As New Dictionary
    params("service") = CStr(service)

    cdp.InvokeMethod "BackgroundService.stopObserving", params
End Sub

' Set the recording state for the service.
Public Sub setRecording( _
    ByVal shouldRecord AS boolean, _
    ByVal service AS string _
)
    ' shouldRecord: boolean 
    ' service: ServiceName 

    Dim params As New Dictionary
    params("shouldRecord") = CBool(shouldRecord)
    params("service") = CStr(service)

    cdp.InvokeMethod "BackgroundService.setRecording", params
End Sub

' Clears all stored data for the service.
Public Sub clearEvents( _
    ByVal service AS string _
)
    ' service: ServiceName 

    Dim params As New Dictionary
    params("service") = CStr(service)

    cdp.InvokeMethod "BackgroundService.clearEvents", params
End Sub

