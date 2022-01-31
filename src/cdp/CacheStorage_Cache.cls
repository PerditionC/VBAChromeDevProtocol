VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCacheStorageCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Cache"
' CacheStorage.Cache
' Cache identifier.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' cacheId : CacheId
' An opaque unique id of the cache.
Public cacheId AS string

' securityOrigin : string
' Security origin of the cache.
Public securityOrigin AS string

' cacheName : string
' The name of the cache.
Public cacheName AS string

Public Function init(ByVal obj as Dictionary) As cdpCacheStorageCache
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let cacheId = CStr(obj.Item("cacheId"))
    Let securityOrigin = CStr(obj.Item("securityOrigin"))
    Let cacheName = CStr(obj.Item("cacheName"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("cacheId") = cacheId
    dict("securityOrigin") = securityOrigin
    dict("cacheName") = cacheName

    Set asDictionary = dict
End Function

