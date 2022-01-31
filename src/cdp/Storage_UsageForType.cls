VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpStorageUsageForType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "UsageForType"
' Storage.UsageForType
' Usage for a storage type.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' storageType : StorageType
' Name of storage type.
Public storageType AS string

' usage : number
' Storage usage (bytes).
Public usage AS Double

Public Function init(ByVal obj as Dictionary) As cdpStorageUsageForType
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let storageType = CStr(obj.Item("storageType"))
    Let usage = CDbl(obj.Item("usage"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("storageType") = storageType
    dict("usage") = usage

    Set asDictionary = dict
End Function

