VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpServiceWorkerServiceWorkerEr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ServiceWorkerErrorMessage"
' ServiceWorker.ServiceWorkerErrorMessage
' ServiceWorker error message.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' errorMessage : string
' 
Public errorMessage AS string

' registrationId : RegistrationID
' 
Public registrationId AS string

' versionId : string
' 
Public versionId AS string

' sourceURL : string
' 
Public sourceURL AS string

' lineNumber : integer
' 
Public lineNumber AS Long

' columnNumber : integer
' 
Public columnNumber AS Long

Public Function init(ByVal obj as Dictionary) As cdpServiceWorkerServiceWorkerEr
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let errorMessage = CStr(obj.Item("errorMessage"))
    Let registrationId = CStr(obj.Item("registrationId"))
    Let versionId = CStr(obj.Item("versionId"))
    Let sourceURL = CStr(obj.Item("sourceURL"))
    Let lineNumber = CLng(obj.Item("lineNumber"))
    Let columnNumber = CLng(obj.Item("columnNumber"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("errorMessage") = errorMessage
    dict("registrationId") = registrationId
    dict("versionId") = versionId
    dict("sourceURL") = sourceURL
    dict("lineNumber") = lineNumber
    dict("columnNumber") = columnNumber

    Set asDictionary = dict
End Function

