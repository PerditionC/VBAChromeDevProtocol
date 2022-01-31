VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimePrivatePropertyDescri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PrivatePropertyDescriptor"
' Runtime.PrivatePropertyDescriptor
' Object private field descriptor.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Private property name.
Public name AS string

' value : RemoteObject
' Optional
' The value associated with the private property.
Public value AS cdpRuntimeRemoteObject

' get : RemoteObject
' Optional
' A function which serves as a getter for the private property,
    '   or `undefined` if there is no getter (accessor descriptors only).
Public getObject AS cdpRuntimeRemoteObject

' set : RemoteObject
' Optional
' A function which serves as a setter for the private property,
    '   or `undefined` if there is no setter (accessor descriptors only).
Public setObject AS cdpRuntimeRemoteObject

Public Function init(ByVal obj as Dictionary) As cdpRuntimePrivatePropertyDescri
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    If obj.Exists("value") Then
        Set value = New cdpRuntimeRemoteObject
        value.init obj.Item("value")
    End If
    If obj.Exists("get") Then
        Set getObject = New cdpRuntimeRemoteObject
        getObject.init obj.Item("get")
    End If
    If obj.Exists("set") Then
        Set setObject = New cdpRuntimeRemoteObject
        setObject.init obj.Item("set")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    If Not value Is Nothing Then Set dict("value") = value.asDictionary()
    If Not getObject Is Nothing Then Set dict("get") = getObject.asDictionary()
    If Not setObject Is Nothing Then Set dict("set") = setObject.asDictionary()

    Set asDictionary = dict
End Function

