VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageInstallabilityErrorArgum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "InstallabilityErrorArgument"
' Page.InstallabilityErrorArgument
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Argument name (e.g. name:'minimum-icon-size-in-pixels').
Public name AS string

' value : string
' Argument value (e.g. value:'64').
Public value AS string

Public Function init(ByVal obj as Dictionary) As cdpPageInstallabilityErrorArgum
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    Let value = CStr(obj.Item("value"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("value") = value

    Set asDictionary = dict
End Function

