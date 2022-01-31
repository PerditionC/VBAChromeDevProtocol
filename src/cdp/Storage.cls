VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpStorage"
' Storage [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Enum of possible storage types.
' StorageType AS string
Public Enum StorageType
    ST_appcache
    ST_cookies
    ST_file_systems
    ST_indexeddb
    ST_local_storage
    ST_shader_cache
    ST_websql
    ST_service_workers
    ST_cache_storage
    ST_all
    ST_other
End Enum

' Usage for a storage type.
' UsageForType AS object

' Pair of issuer origin and number of available (signed, but not used) Trust
'   Tokens from that issuer.
' TrustTokens AS object


Public Function StorageTypeToString(ByVal st As StorageType) As String
    Dim retVal As String
    Select Case st
        Case ST_appcache
            retVal = "appcache"
        Case ST_cookies
            retVal = "cookies"
        Case ST_file_systems
            retVal = "file_systems"
        Case ST_indexeddb
            retVal = "indexeddb"
        Case ST_local_storage
            retVal = "local_storage"
        Case ST_shader_cache
            retVal = "shader_cache"
        Case ST_websql
            retVal = "websql"
        Case ST_service_workers
            retVal = "service_workers"
        Case ST_cache_storage
            retVal = "cache_storage"
        Case ST_all
            retVal = "all"
        Case ST_other
            retVal = "other"
        Case Else
            Debug.Print "Warning, unknown value " & st
    End Select
    StorageTypeToString = retVal
End Function

Public Function StringToStorageType(ByVal s As String) As StorageType
    Dim retVal As StorageType
    Select Case s
        Case "appcache"
            retVal = ST_appcache
        Case "cookies"
            retVal = ST_cookies
        Case "file_systems"
            retVal = ST_file_systems
        Case "indexeddb"
            retVal = ST_indexeddb
        Case "local_storage"
            retVal = ST_local_storage
        Case "shader_cache"
            retVal = ST_shader_cache
        Case "websql"
            retVal = ST_websql
        Case "service_workers"
            retVal = ST_service_workers
        Case "cache_storage"
            retVal = ST_cache_storage
        Case "all"
            retVal = ST_all
        Case "other"
            retVal = ST_other
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToStorageType = retVal
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

' Clears storage for origin.
Public Sub clearDataForOrigin( _
    ByVal origin AS string, _
    ByVal storageTypes AS string _
)
    ' origin: string Security origin.
    ' storageTypes: string Comma separated list of StorageType to clear.

    Dim params As New Dictionary
    params("origin") = CStr(origin)
    params("storageTypes") = CStr(storageTypes)

    cdp.InvokeMethod "Storage.clearDataForOrigin", params
End Sub

' Returns all browser cookies.
' Returns: cookies - Array of cookie objects.
Public Function getCookies( _
    Optional ByVal browserContextId AS Variant _
) AS Collection
    ' browserContextId: Browser.BrowserContextID(optional) Browser context to use when called on the browser endpoint.

    Dim params As New Dictionary
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Storage.getCookies", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("cookies") Then Set getCookies = results("cookies")
    End If
End Function

' Sets given cookies.
Public Sub setCookies( _
    ByVal cookies AS Collection, _
    Optional ByVal browserContextId AS Variant _
)
    ' cookies: array Cookies to be set.
    ' browserContextId: Browser.BrowserContextID(optional) Browser context to use when called on the browser endpoint.

    Dim params As New Dictionary
    Set params("cookies") = cookies
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)

    cdp.InvokeMethod "Storage.setCookies", params
End Sub

' Clears cookies.
Public Sub clearCookies( _
    Optional ByVal browserContextId AS Variant _
)
    ' browserContextId: Browser.BrowserContextID(optional) Browser context to use when called on the browser endpoint.

    Dim params As New Dictionary
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)

    cdp.InvokeMethod "Storage.clearCookies", params
End Sub

' Returns usage and quota in bytes.
' Returns:  - 
Public Function getUsageAndQuota( _
    ByVal origin AS string _
) AS Dictionary
    ' origin: string Security origin.

    Dim params As New Dictionary
    params("origin") = CStr(origin)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Storage.getUsageAndQuota", params)
    If cdp.ErrorCode = 0 Then
        Set getUsageAndQuota = results
    End If
End Function

' Override quota for the specified origin
' Experimental
Public Sub overrideQuotaForOrigin( _
    ByVal origin AS string, _
    Optional ByVal quotaSize AS Variant _
)
    ' origin: string Security origin.
    ' quotaSize: number(optional) The quota size (in bytes) to override the original quota with.
    '   If this is called multiple times, the overridden quota will be equal to
    '   the quotaSize provided in the final call. If this is called without
    '   specifying a quotaSize, the quota will be reset to the default value for
    '   the specified origin. If this is called multiple times with different
    '   origins, the override will be maintained for each origin until it is
    '   disabled (called without a quotaSize).

    Dim params As New Dictionary
    params("origin") = CStr(origin)
    If Not IsMissing(quotaSize) Then params("quotaSize") = CDbl(quotaSize)

    cdp.InvokeMethod "Storage.overrideQuotaForOrigin", params
End Sub

' Registers origin to be notified when an update occurs to its cache storage list.
Public Sub trackCacheStorageForOrigin( _
    ByVal origin AS string _
)
    ' origin: string Security origin.

    Dim params As New Dictionary
    params("origin") = CStr(origin)

    cdp.InvokeMethod "Storage.trackCacheStorageForOrigin", params
End Sub

' Registers origin to be notified when an update occurs to its IndexedDB.
Public Sub trackIndexedDBForOrigin( _
    ByVal origin AS string _
)
    ' origin: string Security origin.

    Dim params As New Dictionary
    params("origin") = CStr(origin)

    cdp.InvokeMethod "Storage.trackIndexedDBForOrigin", params
End Sub

' Unregisters origin from receiving notifications for cache storage.
Public Sub untrackCacheStorageForOrigin( _
    ByVal origin AS string _
)
    ' origin: string Security origin.

    Dim params As New Dictionary
    params("origin") = CStr(origin)

    cdp.InvokeMethod "Storage.untrackCacheStorageForOrigin", params
End Sub

' Unregisters origin from receiving notifications for IndexedDB.
Public Sub untrackIndexedDBForOrigin( _
    ByVal origin AS string _
)
    ' origin: string Security origin.

    Dim params As New Dictionary
    params("origin") = CStr(origin)

    cdp.InvokeMethod "Storage.untrackIndexedDBForOrigin", params
End Sub

' Returns the number of stored Trust Tokens per issuer for the
' current browsing context.
' Returns: tokens - 
' Experimental
Public Function getTrustTokens() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Storage.getTrustTokens", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("tokens") Then Set getTrustTokens = results("tokens")
    End If
End Function

' Removes all Trust Tokens issued by the provided issuerOrigin.
' Leaves other stored data, including the issuer's Redemption Records, intact.
' Returns: didDeleteTokens - True if any tokens were deleted, false otherwise.
' Experimental
Public Function clearTrustTokens( _
    ByVal issuerOrigin AS string _
) AS boolean
    ' issuerOrigin: string 

    Dim params As New Dictionary
    params("issuerOrigin") = CStr(issuerOrigin)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Storage.clearTrustTokens", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("didDeleteTokens") Then clearTrustTokens = results("didDeleteTokens")
    End If
End Function

