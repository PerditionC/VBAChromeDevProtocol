VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpIndexedDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpIndexedDB"
' IndexedDB [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Database with an array of object stores.
' DatabaseWithObjectStores AS object

' Object store.
' ObjectStore AS object

' Object store index.
' ObjectStoreIndex AS object

' Key.
' Key AS object

' Key range.
' KeyRange AS object

' Data entry.
' DataEntry AS object

' Key path.
' KeyPath AS object



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

' Clears all entries from an object store.
Public Sub clearObjectStore( _
    ByVal securityOrigin AS string, _
    ByVal databaseName AS string, _
    ByVal objectStoreName AS string _
)
    ' securityOrigin: string Security origin.
    ' databaseName: string Database name.
    ' objectStoreName: string Object store name.

    Dim params As New Dictionary
    params("securityOrigin") = CStr(securityOrigin)
    params("databaseName") = CStr(databaseName)
    params("objectStoreName") = CStr(objectStoreName)

    cdp.InvokeMethod "IndexedDB.clearObjectStore", params
End Sub

' Deletes a database.
Public Sub deleteDatabase( _
    ByVal securityOrigin AS string, _
    ByVal databaseName AS string _
)
    ' securityOrigin: string Security origin.
    ' databaseName: string Database name.

    Dim params As New Dictionary
    params("securityOrigin") = CStr(securityOrigin)
    params("databaseName") = CStr(databaseName)

    cdp.InvokeMethod "IndexedDB.deleteDatabase", params
End Sub

' Delete a range of entries from an object store
Public Sub deleteObjectStoreEntries( _
    ByVal securityOrigin AS string, _
    ByVal databaseName AS string, _
    ByVal objectStoreName AS string, _
    ByVal keyRange AS Dictionary _
)
    ' securityOrigin: string 
    ' databaseName: string 
    ' objectStoreName: string 
    ' keyRange: KeyRange Range of entry keys to delete

    Dim params As New Dictionary
    params("securityOrigin") = CStr(securityOrigin)
    params("databaseName") = CStr(databaseName)
    params("objectStoreName") = CStr(objectStoreName)
    Set params("keyRange") = keyRange

    cdp.InvokeMethod "IndexedDB.deleteObjectStoreEntries", params
End Sub

' Disables events from backend.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "IndexedDB.disable", params
End Sub

' Enables events from backend.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "IndexedDB.enable", params
End Sub

' Requests data from object store or index.
' Returns:  - 
Public Function requestData( _
    ByVal securityOrigin AS string, _
    ByVal databaseName AS string, _
    ByVal objectStoreName AS string, _
    ByVal indexName AS string, _
    ByVal skipCount AS Long, _
    ByVal pageSize AS Long, _
    Optional ByVal keyRange AS Variant _
) AS Dictionary
    ' securityOrigin: string Security origin.
    ' databaseName: string Database name.
    ' objectStoreName: string Object store name.
    ' indexName: string Index name, empty string for object store data requests.
    ' skipCount: integer Number of records to skip.
    ' pageSize: integer Number of records to fetch.
    ' keyRange: KeyRange(optional) Key range.

    Dim params As New Dictionary
    params("securityOrigin") = CStr(securityOrigin)
    params("databaseName") = CStr(databaseName)
    params("objectStoreName") = CStr(objectStoreName)
    params("indexName") = CStr(indexName)
    params("skipCount") = CLng(skipCount)
    params("pageSize") = CLng(pageSize)
    If Not IsMissing(keyRange) Then Set params("keyRange") = keyRange

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("IndexedDB.requestData", params)
    If cdp.ErrorCode = 0 Then
        Set requestData = results
    End If
End Function

' Gets metadata of an object store
' Returns:  - 
Public Function getMetadata( _
    ByVal securityOrigin AS string, _
    ByVal databaseName AS string, _
    ByVal objectStoreName AS string _
) AS Dictionary
    ' securityOrigin: string Security origin.
    ' databaseName: string Database name.
    ' objectStoreName: string Object store name.

    Dim params As New Dictionary
    params("securityOrigin") = CStr(securityOrigin)
    params("databaseName") = CStr(databaseName)
    params("objectStoreName") = CStr(objectStoreName)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("IndexedDB.getMetadata", params)
    If cdp.ErrorCode = 0 Then
        Set getMetadata = results
    End If
End Function

' Requests database with given name in given frame.
' Returns: databaseWithObjectStores - Database with an array of object stores.
Public Function requestDatabase( _
    ByVal securityOrigin AS string, _
    ByVal databaseName AS string _
) AS Dictionary
    ' securityOrigin: string Security origin.
    ' databaseName: string Database name.

    Dim params As New Dictionary
    params("securityOrigin") = CStr(securityOrigin)
    params("databaseName") = CStr(databaseName)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("IndexedDB.requestDatabase", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("databaseWithObjectStores") Then Set requestDatabase = results("databaseWithObjectStores")
    End If
End Function

' Requests database names for given security origin.
' Returns: databaseNames - Database names for origin.
Public Function requestDatabaseNames( _
    ByVal securityOrigin AS string _
) AS Collection
    ' securityOrigin: string Security origin.

    Dim params As New Dictionary
    params("securityOrigin") = CStr(securityOrigin)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("IndexedDB.requestDatabaseNames", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("databaseNames") Then Set requestDatabaseNames = results("databaseNames")
    End If
End Function

