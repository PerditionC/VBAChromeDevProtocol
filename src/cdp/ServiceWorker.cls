VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpServiceWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpServiceWorker"
' ServiceWorker [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' 
' RegistrationID AS string

' ServiceWorker registration.
' ServiceWorkerRegistration AS object

' 
' ServiceWorkerVersionRunningStatus AS string
Public Enum ServiceWorkerVersionRunningStatus
    SWVRS_stopped
    SWVRS_starting
    SWVRS_running
    SWVRS_stopping
End Enum

' 
' ServiceWorkerVersionStatus AS string
Public Enum ServiceWorkerVersionStatus
    SWVS_new
    SWVS_installing
    SWVS_installed
    SWVS_activating
    SWVS_activated
    SWVS_redundant
End Enum

' ServiceWorker version.
' ServiceWorkerVersion AS object

' ServiceWorker error message.
' ServiceWorkerErrorMessage AS object


Public Function ServiceWorkerVersionRunningStatusToString(ByVal swvrs As ServiceWorkerVersionRunningStatus) As String
    Dim retVal As String
    Select Case swvrs
        Case SWVRS_stopped
            retVal = "stopped"
        Case SWVRS_starting
            retVal = "starting"
        Case SWVRS_running
            retVal = "running"
        Case SWVRS_stopping
            retVal = "stopping"
        Case Else
            Debug.Print "Warning, unknown value " & swvrs
    End Select
    ServiceWorkerVersionRunningStatusToString = retVal
End Function

Public Function StringToServiceWorkerVersionRunningStatus(ByVal s As String) As ServiceWorkerVersionRunningStatus
    Dim retVal As ServiceWorkerVersionRunningStatus
    Select Case s
        Case "stopped"
            retVal = SWVRS_stopped
        Case "starting"
            retVal = SWVRS_starting
        Case "running"
            retVal = SWVRS_running
        Case "stopping"
            retVal = SWVRS_stopping
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToServiceWorkerVersionRunningStatus = retVal
End Function


Public Function ServiceWorkerVersionStatusToString(ByVal swvs As ServiceWorkerVersionStatus) As String
    Dim retVal As String
    Select Case swvs
        Case SWVS_new
            retVal = "new"
        Case SWVS_installing
            retVal = "installing"
        Case SWVS_installed
            retVal = "installed"
        Case SWVS_activating
            retVal = "activating"
        Case SWVS_activated
            retVal = "activated"
        Case SWVS_redundant
            retVal = "redundant"
        Case Else
            Debug.Print "Warning, unknown value " & swvs
    End Select
    ServiceWorkerVersionStatusToString = retVal
End Function

Public Function StringToServiceWorkerVersionStatus(ByVal s As String) As ServiceWorkerVersionStatus
    Dim retVal As ServiceWorkerVersionStatus
    Select Case s
        Case "new"
            retVal = SWVS_new
        Case "installing"
            retVal = SWVS_installing
        Case "installed"
            retVal = SWVS_installed
        Case "activating"
            retVal = SWVS_activating
        Case "activated"
            retVal = SWVS_activated
        Case "redundant"
            retVal = SWVS_redundant
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToServiceWorkerVersionStatus = retVal
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

' 
Public Sub deliverPushMessage( _
    ByVal origin AS string, _
    ByVal registrationId AS string, _
    ByVal data AS string _
)
    ' origin: string 
    ' registrationId: RegistrationID 
    ' data: string 

    Dim params As New Dictionary
    params("origin") = CStr(origin)
    params("registrationId") = CStr(registrationId)
    params("data") = CStr(data)

    cdp.InvokeMethod "ServiceWorker.deliverPushMessage", params
End Sub

' 
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "ServiceWorker.disable", params
End Sub

' 
Public Sub dispatchSyncEvent( _
    ByVal origin AS string, _
    ByVal registrationId AS string, _
    ByVal tag AS string, _
    ByVal lastChance AS boolean _
)
    ' origin: string 
    ' registrationId: RegistrationID 
    ' tag: string 
    ' lastChance: boolean 

    Dim params As New Dictionary
    params("origin") = CStr(origin)
    params("registrationId") = CStr(registrationId)
    params("tag") = CStr(tag)
    params("lastChance") = CBool(lastChance)

    cdp.InvokeMethod "ServiceWorker.dispatchSyncEvent", params
End Sub

' 
Public Sub dispatchPeriodicSyncEvent( _
    ByVal origin AS string, _
    ByVal registrationId AS string, _
    ByVal tag AS string _
)
    ' origin: string 
    ' registrationId: RegistrationID 
    ' tag: string 

    Dim params As New Dictionary
    params("origin") = CStr(origin)
    params("registrationId") = CStr(registrationId)
    params("tag") = CStr(tag)

    cdp.InvokeMethod "ServiceWorker.dispatchPeriodicSyncEvent", params
End Sub

' 
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "ServiceWorker.enable", params
End Sub

' 
Public Sub inspectWorker( _
    ByVal versionId AS string _
)
    ' versionId: string 

    Dim params As New Dictionary
    params("versionId") = CStr(versionId)

    cdp.InvokeMethod "ServiceWorker.inspectWorker", params
End Sub

' 
Public Sub setForceUpdateOnPageLoad( _
    ByVal forceUpdateOnPageLoad AS boolean _
)
    ' forceUpdateOnPageLoad: boolean 

    Dim params As New Dictionary
    params("forceUpdateOnPageLoad") = CBool(forceUpdateOnPageLoad)

    cdp.InvokeMethod "ServiceWorker.setForceUpdateOnPageLoad", params
End Sub

' 
Public Sub skipWaiting( _
    ByVal scopeURL AS string _
)
    ' scopeURL: string 

    Dim params As New Dictionary
    params("scopeURL") = CStr(scopeURL)

    cdp.InvokeMethod "ServiceWorker.skipWaiting", params
End Sub

' 
Public Sub startWorker( _
    ByVal scopeURL AS string _
)
    ' scopeURL: string 

    Dim params As New Dictionary
    params("scopeURL") = CStr(scopeURL)

    cdp.InvokeMethod "ServiceWorker.startWorker", params
End Sub

' 
Public Sub stopAllWorkers()
    Dim params As New Dictionary
    cdp.InvokeMethod "ServiceWorker.stopAllWorkers", params
End Sub

' 
Public Sub stopWorker( _
    ByVal versionId AS string _
)
    ' versionId: string 

    Dim params As New Dictionary
    params("versionId") = CStr(versionId)

    cdp.InvokeMethod "ServiceWorker.stopWorker", params
End Sub

' 
Public Sub unregister( _
    ByVal scopeURL AS string _
)
    ' scopeURL: string 

    Dim params As New Dictionary
    params("scopeURL") = CStr(scopeURL)

    cdp.InvokeMethod "ServiceWorker.unregister", params
End Sub

' 
Public Sub updateRegistration( _
    ByVal scopeURL AS string _
)
    ' scopeURL: string 

    Dim params As New Dictionary
    params("scopeURL") = CStr(scopeURL)

    cdp.InvokeMethod "ServiceWorker.updateRegistration", params
End Sub

