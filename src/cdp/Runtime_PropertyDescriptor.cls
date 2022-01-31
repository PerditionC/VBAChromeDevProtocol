VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimePropertyDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PropertyDescriptor"
' Runtime.PropertyDescriptor
' Object property descriptor.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Property name or symbol description.
Public name AS string

' value : RemoteObject
' Optional
' The value associated with the property.
Public value AS cdpRuntimeRemoteObject

' writable : boolean
' Optional
' True if the value associated with the property may be changed (data descriptors only).
Public writable AS Variant ' boolean

' get : RemoteObject
' Optional
' A function which serves as a getter for the property, or `undefined` if there is no getter
    '   (accessor descriptors only).
Public getObject AS cdpRuntimeRemoteObject

' set : RemoteObject
' Optional
' A function which serves as a setter for the property, or `undefined` if there is no setter
    '   (accessor descriptors only).
Public setObject AS cdpRuntimeRemoteObject

' configurable : boolean
' True if the type of this property descriptor may be changed and if the property may be
    '   deleted from the corresponding object.
Public configurable AS boolean

' enumerable : boolean
' True if this property shows up during enumeration of the properties on the corresponding
    '   object.
Public enumerable AS boolean

' wasThrown : boolean
' Optional
' True if the result was thrown during the evaluation.
Public wasThrown AS Variant ' boolean

' isOwn : boolean
' Optional
' True if the property is owned for the object.
Public isOwn AS Variant ' boolean

' symbol : RemoteObject
' Optional
' Property symbol object, if the property is of the `symbol` type.
Public symbol AS cdpRuntimeRemoteObject

Public Function init(ByVal obj as Dictionary) As cdpRuntimePropertyDescriptor
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    If obj.Exists("value") Then
        Set value = New cdpRuntimeRemoteObject
        value.init obj.Item("value")
    End If
    If obj.Exists("writable") Then Let writable = CBool(obj.Item("writable"))
    If obj.Exists("get") Then
        Set getObject = New cdpRuntimeRemoteObject
        getObject.init obj.Item("get")
    End If
    If obj.Exists("set") Then
        Set setObject = New cdpRuntimeRemoteObject
        setObject.init obj.Item("set")
    End If
    Let configurable = CBool(obj.Item("configurable"))
    Let enumerable = CBool(obj.Item("enumerable"))
    If obj.Exists("wasThrown") Then Let wasThrown = CBool(obj.Item("wasThrown"))
    If obj.Exists("isOwn") Then Let isOwn = CBool(obj.Item("isOwn"))
    If obj.Exists("symbol") Then
        Set symbol = New cdpRuntimeRemoteObject
        symbol.init obj.Item("symbol")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    If Not value Is Nothing Then Set dict("value") = value.asDictionary()
    if Not IsEmpty(writable) Then dict("writable") = writable
    If Not getObject Is Nothing Then Set dict("get") = getObject.asDictionary()
    If Not setObject Is Nothing Then Set dict("set") = setObject.asDictionary()
    dict("configurable") = configurable
    dict("enumerable") = enumerable
    if Not IsEmpty(wasThrown) Then dict("wasThrown") = wasThrown
    if Not IsEmpty(isOwn) Then dict("isOwn") = isOwn
    If Not symbol Is Nothing Then Set dict("symbol") = symbol.asDictionary()

    Set asDictionary = dict
End Function

