VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCacheStorageCachedResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CachedResponse"
' CacheStorage.CachedResponse
' Cached response
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' body : binary
' Entry content, base64-encoded.
Public body AS String

Public Function init(ByVal obj as Dictionary) As cdpCacheStorageCachedResponse
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let body = CStr(obj.Item("body"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("body") = body

    Set asDictionary = dict
End Function

