VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCacheStorageDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DataEntry"
' CacheStorage.DataEntry
' Data entry.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' requestURL : string
' Request URL.
Public requestURL AS string

' requestMethod : string
' Request method.
Public requestMethod AS string

' requestHeaders : array of cdpCacheStorageHeader
' Request headers
Public requestHeaders AS Collection

' responseTime : number
' Number of seconds since epoch.
Public responseTime AS Double

' responseStatus : integer
' HTTP response status code.
Public responseStatus AS Long

' responseStatusText : string
' HTTP response status text.
Public responseStatusText AS string

' responseType : CachedResponseType
' HTTP response type
Public responseType AS string

' responseHeaders : array of cdpCacheStorageHeader
' Response headers
Public responseHeaders AS Collection

Public Function init(ByVal obj as Dictionary) As cdpCacheStorageDataEntry
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let requestURL = CStr(obj.Item("requestURL"))
    Let requestMethod = CStr(obj.Item("requestMethod"))
    For each v in obj.Item("requestHeaders")
        Dim obj_requestHeaders As cdpCacheStorageHeader: Set obj_requestHeaders = New cdpCacheStorageHeader
        obj_requestHeaders.init v
        requestHeaders.Add obj_requestHeaders
        Set obj_requestHeaders = Nothing
    Next v
    Let responseTime = CDbl(obj.Item("responseTime"))
    Let responseStatus = CLng(obj.Item("responseStatus"))
    Let responseStatusText = CStr(obj.Item("responseStatusText"))
    Let responseType = CStr(obj.Item("responseType"))
    For each v in obj.Item("responseHeaders")
        Dim obj_responseHeaders As cdpCacheStorageHeader: Set obj_responseHeaders = New cdpCacheStorageHeader
        obj_responseHeaders.init v
        responseHeaders.Add obj_responseHeaders
        Set obj_responseHeaders = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("requestURL") = requestURL
    dict("requestMethod") = requestMethod
    Set dict("requestHeaders") = requestHeaders
    dict("responseTime") = responseTime
    dict("responseStatus") = responseStatus
    dict("responseStatusText") = responseStatusText
    dict("responseType") = responseType
    Set dict("responseHeaders") = responseHeaders

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set requestHeaders = New Collection
    Set responseHeaders = New Collection
End Sub

