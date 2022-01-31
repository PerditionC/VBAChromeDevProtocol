VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMStorageStorageId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "StorageId"
' DOMStorage.StorageId
' DOM Storage identifier.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' securityOrigin : string
' Security origin for the storage.
Public securityOrigin AS string

' isLocalStorage : boolean
' Whether the storage is local storage (not session storage).
Public isLocalStorage AS boolean

Public Function init(ByVal obj as Dictionary) As cdpDOMStorageStorageId
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let securityOrigin = CStr(obj.Item("securityOrigin"))
    Let isLocalStorage = CBool(obj.Item("isLocalStorage"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("securityOrigin") = securityOrigin
    dict("isLocalStorage") = isLocalStorage

    Set asDictionary = dict
End Function

