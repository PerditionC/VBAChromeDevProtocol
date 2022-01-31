VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDatabaseDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Database"
' Database.Database
' Database object.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' id : DatabaseId
' Database ID.
Public id AS string

' domain : string
' Database domain.
Public domain AS string

' name : string
' Database name.
Public name AS string

' version : string
' Database version.
Public version AS string

Public Function init(ByVal obj as Dictionary) As cdpDatabaseDatabase
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let id = CStr(obj.Item("id"))
    Let domain = CStr(obj.Item("domain"))
    Let name = CStr(obj.Item("name"))
    Let version = CStr(obj.Item("version"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("id") = id
    dict("domain") = domain
    dict("name") = name
    dict("version") = version

    Set asDictionary = dict
End Function

