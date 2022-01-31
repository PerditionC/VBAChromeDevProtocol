VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpDatabase"
' Database [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Unique identifier of Database object.
' DatabaseId AS string

' Database object.
' Database AS object

' Database error.
' Error AS object



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

' Disables database tracking, prevents database events from being sent to the client.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Database.disable", params
End Sub

' Enables database tracking, database events will now be delivered to the client.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Database.enable", params
End Sub

' 
' Returns:  - 
Public Function executeSQL( _
    ByVal databaseId AS string, _
    ByVal query AS string _
) AS Dictionary
    ' databaseId: DatabaseId 
    ' query: string 

    Dim params As New Dictionary
    params("databaseId") = CStr(databaseId)
    params("query") = CStr(query)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Database.executeSQL", params)
    If cdp.ErrorCode = 0 Then
        Set executeSQL = results
    End If
End Function

' 
' Returns: tableNames - 
Public Function getDatabaseTableNames( _
    ByVal databaseId AS string _
) AS Collection
    ' databaseId: DatabaseId 

    Dim params As New Dictionary
    params("databaseId") = CStr(databaseId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Database.getDatabaseTableNames", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("tableNames") Then Set getDatabaseTableNames = results("tableNames")
    End If
End Function

