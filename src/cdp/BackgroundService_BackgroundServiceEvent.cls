VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBackgroundServiceBackgroundS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "BackgroundServiceEvent"
' BackgroundService.BackgroundServiceEvent
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' timestamp : Network.TimeSinceEpoch
' Timestamp of the event (in seconds).
Public timestamp AS Double

' origin : string
' The origin this event belongs to.
Public origin AS string

' serviceWorkerRegistrationId : ServiceWorker.RegistrationID
' The Service Worker ID that initiated the event.
Public serviceWorkerRegistrationId AS string

' service : ServiceName
' The Background Service this event belongs to.
Public service AS string

' eventName : string
' A description of the event.
Public eventName AS string

' instanceId : string
' An identifier that groups related events together.
Public instanceId AS string

' eventMetadata : array of cdpBackgroundServiceEventMetada
' A list of event-specific information.
Public eventMetadata AS Collection

Public Function init(ByVal obj as Dictionary) As cdpBackgroundServiceBackgroundS
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let timestamp = CDbl(obj.Item("timestamp"))
    Let origin = CStr(obj.Item("origin"))
    Let serviceWorkerRegistrationId = CStr(obj.Item("serviceWorkerRegistrationId"))
    Let service = CStr(obj.Item("service"))
    Let eventName = CStr(obj.Item("eventName"))
    Let instanceId = CStr(obj.Item("instanceId"))
    For each v in obj.Item("eventMetadata")
        Dim obj_eventMetadata As cdpBackgroundServiceEventMetada: Set obj_eventMetadata = New cdpBackgroundServiceEventMetada
        obj_eventMetadata.init v
        eventMetadata.Add obj_eventMetadata
        Set obj_eventMetadata = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("timestamp") = timestamp
    dict("origin") = origin
    dict("serviceWorkerRegistrationId") = serviceWorkerRegistrationId
    dict("service") = service
    dict("eventName") = eventName
    dict("instanceId") = instanceId
    Set dict("eventMetadata") = eventMetadata

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set eventMetadata = New Collection
End Sub

