VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpDOMStorage"
' DOMStorage [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' DOM Storage identifier.
' StorageId AS object

' DOM Storage item.
' Item AS array



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

' 
Public Sub clear( _
    ByVal storageId AS Dictionary _
)
    ' storageId: StorageId 

    Dim params As New Dictionary
    Set params("storageId") = storageId

    cdp.InvokeMethod "DOMStorage.clear", params
End Sub

' Disables storage tracking, prevents storage events from being sent to the client.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOMStorage.disable", params
End Sub

' Enables storage tracking, storage events will now be delivered to the client.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "DOMStorage.enable", params
End Sub

' 
' Returns: entries - 
Public Function getDOMStorageItems( _
    ByVal storageId AS Dictionary _
) AS Collection
    ' storageId: StorageId 

    Dim params As New Dictionary
    Set params("storageId") = storageId

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("DOMStorage.getDOMStorageItems", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("entries") Then Set getDOMStorageItems = results("entries")
    End If
End Function

' 
Public Sub removeDOMStorageItem( _
    ByVal storageId AS Dictionary, _
    ByVal key AS string _
)
    ' storageId: StorageId 
    ' key: string 

    Dim params As New Dictionary
    Set params("storageId") = storageId
    params("key") = CStr(key)

    cdp.InvokeMethod "DOMStorage.removeDOMStorageItem", params
End Sub

' 
Public Sub setDOMStorageItem( _
    ByVal storageId AS Dictionary, _
    ByVal key AS string, _
    ByVal value AS string _
)
    ' storageId: StorageId 
    ' key: string 
    ' value: string 

    Dim params As New Dictionary
    Set params("storageId") = storageId
    params("key") = CStr(key)
    params("value") = CStr(value)

    cdp.InvokeMethod "DOMStorage.setDOMStorageItem", params
End Sub

