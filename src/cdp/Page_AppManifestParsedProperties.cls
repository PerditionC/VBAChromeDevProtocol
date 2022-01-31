VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageAppManifestParsedPropert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AppManifestParsedProperties"
' Page.AppManifestParsedProperties
' Parsed app manifest properties.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' scope : string
' Computed scope value
Public scope AS string

Public Function init(ByVal obj as Dictionary) As cdpPageAppManifestParsedPropert
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let scope = CStr(obj.Item("scope"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("scope") = scope

    Set asDictionary = dict
End Function

