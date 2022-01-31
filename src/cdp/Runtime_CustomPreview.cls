VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeCustomPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CustomPreview"
' Runtime.CustomPreview
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' header : string
' The JSON-stringified result of formatter.header(object, config) call.
    '   It contains json ML array that represents RemoteObject.
Public header AS string

' bodyGetterId : RemoteObjectId
' Optional
' If formatter returns true as a result of formatter.hasBody call then bodyGetterId will
    '   contain RemoteObjectId for the function that returns result of formatter.body(object, config) call.
    '   The result value is json ML array.
Public bodyGetterId AS string

Public Function init(ByVal obj as Dictionary) As cdpRuntimeCustomPreview
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let header = CStr(obj.Item("header"))
    If obj.Exists("bodyGetterId") Then Let bodyGetterId = CStr(obj.Item("bodyGetterId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("header") = header
    if bodyGetterId <> vbNullString Then dict("bodyGetterId") = bodyGetterId

    Set asDictionary = dict
End Function

