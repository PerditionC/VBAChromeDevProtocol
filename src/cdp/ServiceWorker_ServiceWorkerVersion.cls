VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpServiceWorkerServiceWorkerVe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ServiceWorkerVersion"
' ServiceWorker.ServiceWorkerVersion
' ServiceWorker version.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' versionId : string
' 
Public versionId AS string

' registrationId : RegistrationID
' 
Public registrationId AS string

' scriptURL : string
' 
Public scriptURL AS string

' runningStatus : ServiceWorkerVersionRunningStatus
' 
Public runningStatus AS string

' status : ServiceWorkerVersionStatus
' 
Public status AS string

' scriptLastModified : number
' Optional
' The Last-Modified header value of the main script.
Public scriptLastModified AS Double

' scriptResponseTime : number
' Optional
' The time at which the response headers of the main script were received from the server.
    '   For cached script it is the last time the cache entry was validated.
Public scriptResponseTime AS Double

' controlledClients : array of cdpTargetTargetID
' Optional
' 
Public controlledClients AS Collection

' targetId : Target.TargetID
' Optional
' 
Public targetId AS string

Public Function init(ByVal obj as Dictionary) As cdpServiceWorkerServiceWorkerVe
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let versionId = CStr(obj.Item("versionId"))
    Let registrationId = CStr(obj.Item("registrationId"))
    Let scriptURL = CStr(obj.Item("scriptURL"))
    Let runningStatus = CStr(obj.Item("runningStatus"))
    Let status = CStr(obj.Item("status"))
    If obj.Exists("scriptLastModified") Then Let scriptLastModified = CDbl(obj.Item("scriptLastModified"))
    If obj.Exists("scriptResponseTime") Then Let scriptResponseTime = CDbl(obj.Item("scriptResponseTime"))
    If obj.Exists("controlledClients") Then
        For each v in obj.Item("controlledClients")
            controlledClients.Add v
        Next v
    End If
    If obj.Exists("targetId") Then Let targetId = CStr(obj.Item("targetId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("versionId") = versionId
    dict("registrationId") = registrationId
    dict("scriptURL") = scriptURL
    dict("runningStatus") = runningStatus
    dict("status") = status
    dict("scriptLastModified") = scriptLastModified
    dict("scriptResponseTime") = scriptResponseTime
    Set dict("controlledClients") = controlledClients
    if targetId <> vbNullString Then dict("targetId") = targetId

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set controlledClients = New Collection
End Sub

