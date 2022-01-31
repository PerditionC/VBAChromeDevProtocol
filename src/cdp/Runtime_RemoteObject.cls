VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeRemoteObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "RemoteObject"
' Runtime.RemoteObject
' Mirror object referencing original JavaScript object.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : string
' Object type.
Public fieldType AS string

' subtype : string
' Optional
' Object subtype hint. Specified for `object` type values only.
    '   NOTE: If you change anything here, make sure to also update
    '   `subtype` in `ObjectPreview` and `PropertyPreview` below.
Public subtype AS string

' className : string
' Optional
' Object class (constructor) name. Specified for `object` type values only.
Public className AS string

' value : any
' Optional
' Remote object value in case of primitive values or JSON values (if it was requested).
Public value AS Variant

' unserializableValue : UnserializableValue
' Optional
' Primitive value which can not be JSON-stringified does not have `value`, but gets this
    '   property.
Public unserializableValue AS string

' description : string
' Optional
' String representation of the object.
Public description AS string

' objectId : RemoteObjectId
' Optional
' Unique object identifier (for non-primitive values).
Public objectId AS string

' preview : ObjectPreview
' Optional
' Preview containing abbreviated property values. Specified for `object` type values only.
Public preview AS cdpRuntimeObjectPreview

' customPreview : CustomPreview
' Optional
' 
Public customPreview AS cdpRuntimeCustomPreview

Public Function init(ByVal obj as Dictionary) As cdpRuntimeRemoteObject
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let fieldType = CStr(obj.Item("type"))
    If obj.Exists("subtype") Then Let subtype = CStr(obj.Item("subtype"))
    If obj.Exists("className") Then Let className = CStr(obj.Item("className"))
    If obj.Exists("value") Then
        If IsObject(obj.Item("value")) Then
            Set value = obj.Item("value")
        Else
            Let value = obj.Item("value")
        End If
    End If
    If obj.Exists("unserializableValue") Then Let unserializableValue = CStr(obj.Item("unserializableValue"))
    If obj.Exists("description") Then Let description = CStr(obj.Item("description"))
    If obj.Exists("objectId") Then Let objectId = CStr(obj.Item("objectId"))
    If obj.Exists("preview") Then
        Set preview = New cdpRuntimeObjectPreview
        preview.init obj.Item("preview")
    End If
    If obj.Exists("customPreview") Then
        Set customPreview = New cdpRuntimeCustomPreview
        customPreview.init obj.Item("customPreview")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    if subtype <> vbNullString Then dict("subtype") = subtype
    if className <> vbNullString Then dict("className") = className
    dict("value") = value
    if unserializableValue <> vbNullString Then dict("unserializableValue") = unserializableValue
    if description <> vbNullString Then dict("description") = description
    if objectId <> vbNullString Then dict("objectId") = objectId
    If Not preview Is Nothing Then Set dict("preview") = preview.asDictionary()
    If Not customPreview Is Nothing Then Set dict("customPreview") = customPreview.asDictionary()

    Set asDictionary = dict
End Function

