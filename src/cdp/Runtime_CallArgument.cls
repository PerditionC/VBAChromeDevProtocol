VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeCallArgument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CallArgument"
' Runtime.CallArgument
' Represents function call argument. Either remote object id `objectId`, primitive `value`,
'   unserializable primitive value or neither of (for undefined) them should be specified.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' value : any
' Optional
' Primitive value or serializable javascript object.
Public value AS Variant

' unserializableValue : UnserializableValue
' Optional
' Primitive value which can not be JSON-stringified.
Public unserializableValue AS string

' objectId : RemoteObjectId
' Optional
' Remote object handle.
Public objectId AS string

Public Function init(ByVal obj as Dictionary) As cdpRuntimeCallArgument
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("value") Then
        If IsObject(obj.Item("value")) Then
            Set value = obj.Item("value")
        Else
            Let value = obj.Item("value")
        End If
    End If
    If obj.Exists("unserializableValue") Then Let unserializableValue = CStr(obj.Item("unserializableValue"))
    If obj.Exists("objectId") Then Let objectId = CStr(obj.Item("objectId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("value") = value
    if unserializableValue <> vbNullString Then dict("unserializableValue") = unserializableValue
    if objectId <> vbNullString Then dict("objectId") = objectId

    Set asDictionary = dict
End Function

