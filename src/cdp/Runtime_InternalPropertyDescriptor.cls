VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeInternalPropertyDescr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "InternalPropertyDescriptor"
' Runtime.InternalPropertyDescriptor
' Object internal property descriptor. This property isn't normally visible in JavaScript code.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Conventional property name.
Public name AS string

' value : RemoteObject
' Optional
' The value associated with the property.
Public value AS cdpRuntimeRemoteObject

Public Function init(ByVal obj as Dictionary) As cdpRuntimeInternalPropertyDescr
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    If obj.Exists("value") Then
        Set value = New cdpRuntimeRemoteObject
        value.init obj.Item("value")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    If Not value Is Nothing Then Set dict("value") = value.asDictionary()

    Set asDictionary = dict
End Function

