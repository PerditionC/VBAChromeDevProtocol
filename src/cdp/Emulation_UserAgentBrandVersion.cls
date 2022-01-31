VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpEmulationUserAgentBrandVersi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "UserAgentBrandVersion"
' Emulation.UserAgentBrandVersion
' Used to specify User Agent Cient Hints to emulate. See https://wicg.github.io/ua-client-hints
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' brand : string
' 
Public brand AS string

' version : string
' 
Public version AS string

Public Function init(ByVal obj as Dictionary) As cdpEmulationUserAgentBrandVersi
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let brand = CStr(obj.Item("brand"))
    Let version = CStr(obj.Item("version"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("brand") = brand
    dict("version") = version

    Set asDictionary = dict
End Function

