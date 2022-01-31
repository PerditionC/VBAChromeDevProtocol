VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpServiceWorkerServiceWorkerRe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ServiceWorkerRegistration"
' ServiceWorker.ServiceWorkerRegistration
' ServiceWorker registration.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' registrationId : RegistrationID
' 
Public registrationId AS string

' scopeURL : string
' 
Public scopeURL AS string

' isDeleted : boolean
' 
Public isDeleted AS boolean

Public Function init(ByVal obj as Dictionary) As cdpServiceWorkerServiceWorkerRe
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let registrationId = CStr(obj.Item("registrationId"))
    Let scopeURL = CStr(obj.Item("scopeURL"))
    Let isDeleted = CBool(obj.Item("isDeleted"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("registrationId") = registrationId
    dict("scopeURL") = scopeURL
    dict("isDeleted") = isDeleted

    Set asDictionary = dict
End Function

