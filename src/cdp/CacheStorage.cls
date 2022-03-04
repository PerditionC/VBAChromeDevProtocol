VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCacheStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpCacheStorage"
' CacheStorage [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Unique identifier of the Cache object.
' CacheId AS string

' type of HTTP response cached
' CachedResponseType AS string
Public Enum CachedResponseType
    CRT_basic
    CRT_cors
    CRT_default
    CRT_error
    CRT_opaqueResponse
    CRT_opaqueRedirect
End Enum

' Data entry.
' DataEntry AS object

' Cache identifier.
' Cache AS object

' 
' Header AS object

' Cached response
' CachedResponse AS object


Public Function CachedResponseTypeToString(ByVal crt As CachedResponseType) As String
    Dim retVal As String
    Select Case crt
        Case CRT_basic
            retVal = "basic"
        Case CRT_cors
            retVal = "cors"
        Case CRT_default
            retVal = "default"
        Case CRT_error
            retVal = "error"
        Case CRT_opaqueResponse
            retVal = "opaqueResponse"
        Case CRT_opaqueRedirect
            retVal = "opaqueRedirect"
        Case Else
            Debug.Print "Warning, unknown value " & crt
    End Select
    CachedResponseTypeToString = retVal
End Function

Public Function StringToCachedResponseType(ByVal s As String) As CachedResponseType
    Dim retVal As CachedResponseType
    Select Case s
        Case "basic"
            retVal = CRT_basic
        Case "cors"
            retVal = CRT_cors
        Case "default"
            retVal = CRT_default
        Case "error"
            retVal = CRT_error
        Case "opaqueResponse"
            retVal = CRT_opaqueResponse
        Case "opaqueRedirect"
            retVal = CRT_opaqueRedirect
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCachedResponseType = retVal
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

' Deletes a cache.
Public Sub deleteCache( _
    ByVal cacheId AS string _
)
    ' cacheId: CacheId Id of cache for deletion.

    Dim params As New Dictionary
    params("cacheId") = CStr(cacheId)

    cdp.InvokeMethod "CacheStorage.deleteCache", params
End Sub

' Deletes a cache entry.
Public Sub deleteEntry( _
    ByVal cacheId AS string, _
    ByVal request AS string _
)
    ' cacheId: CacheId Id of cache where the entry will be deleted.
    ' request: string URL spec of the request.

    Dim params As New Dictionary
    params("cacheId") = CStr(cacheId)
    params("request") = CStr(request)

    cdp.InvokeMethod "CacheStorage.deleteEntry", params
End Sub

' Requests cache names.
' Returns: caches - Caches for the security origin.
Public Function requestCacheNames( _
    ByVal securityOrigin AS string _
) AS Collection
    ' securityOrigin: string Security origin.

    Dim params As New Dictionary
    params("securityOrigin") = CStr(securityOrigin)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CacheStorage.requestCacheNames", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("caches") Then Set requestCacheNames = results("caches")
    End If
End Function

' Fetches cache entry.
' Returns: response - Response read from the cache.
Public Function requestCachedResponse( _
    ByVal cacheId AS string, _
    ByVal requestURL AS string, _
    ByVal requestHeaders AS Collection _
) AS Dictionary
    ' cacheId: CacheId Id of cache that contains the entry.
    ' requestURL: string URL spec of the request.
    ' requestHeaders: array headers of the request.

    Dim params As New Dictionary
    params("cacheId") = CStr(cacheId)
    params("requestURL") = CStr(requestURL)
    Set params("requestHeaders") = requestHeaders

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CacheStorage.requestCachedResponse", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("response") Then Set requestCachedResponse = results("response")
    End If
End Function

' Requests data from cache.
' Returns:  - 
Public Function requestEntries( _
    ByVal cacheId AS string, _
    Optional ByVal skipCount AS Variant, _
    Optional ByVal pageSize AS Variant, _
    Optional ByVal pathFilter AS Variant _
) AS Dictionary
    ' cacheId: CacheId ID of cache to get entries from.
    ' skipCount: integer(optional) Number of records to skip.
    ' pageSize: integer(optional) Number of records to fetch.
    ' pathFilter: string(optional) If present, only return the entries containing this substring in the path

    Dim params As New Dictionary
    params("cacheId") = CStr(cacheId)
    If Not IsMissing(skipCount) Then params("skipCount") = CLng(skipCount)
    If Not IsMissing(pageSize) Then params("pageSize") = CLng(pageSize)
    If Not IsMissing(pathFilter) Then params("pathFilter") = CStr(pathFilter)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CacheStorage.requestEntries", params)
    If cdp.ErrorCode = 0 Then
        Set requestEntries = results
    End If
End Function

