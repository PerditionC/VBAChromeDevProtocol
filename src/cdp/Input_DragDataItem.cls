VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpInputDragDataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DragDataItem"
' Input.DragDataItem
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' mimeType : string
' Mime type of the dragged data.
Public mimeType AS string

' data : string
' Depending of the value of `mimeType`, it contains the dragged link,
    '   text, HTML markup or any other data.
Public data AS string

' title : string
' Optional
' Title associated with a link. Only valid when `mimeType` == "text/uri-list".
Public title AS string

' baseURL : string
' Optional
' Stores the base URL for the contained markup. Only valid when `mimeType`
    '   == "text/html".
Public baseURL AS string

Public Function init(ByVal obj as Dictionary) As cdpInputDragDataItem
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let mimeType = CStr(obj.Item("mimeType"))
    Let data = CStr(obj.Item("data"))
    If obj.Exists("title") Then Let title = CStr(obj.Item("title"))
    If obj.Exists("baseURL") Then Let baseURL = CStr(obj.Item("baseURL"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("mimeType") = mimeType
    dict("data") = data
    if title <> vbNullString Then dict("title") = title
    if baseURL <> vbNullString Then dict("baseURL") = baseURL

    Set asDictionary = dict
End Function

