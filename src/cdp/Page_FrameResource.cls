VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageFrameResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FrameResource"
' Page.FrameResource
' Information about the Resource on the page.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' url : string
' Resource URL.
Public url AS string

' type : Network.ResourceType
' Type of this resource.
Public fieldType AS string

' mimeType : string
' Resource mimeType as determined by the browser.
Public mimeType AS string

' lastModified : Network.TimeSinceEpoch
' Optional
' last-modified timestamp as reported by server.
Public lastModified AS Double

' contentSize : number
' Optional
' Resource content size.
Public contentSize AS Double

' failed : boolean
' Optional
' True if the resource failed to load.
Public failed AS Variant ' boolean

' canceled : boolean
' Optional
' True if the resource was canceled during loading.
Public canceled AS Variant ' boolean

Public Function init(ByVal obj as Dictionary) As cdpPageFrameResource
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let url = CStr(obj.Item("url"))
    Let fieldType = CStr(obj.Item("type"))
    Let mimeType = CStr(obj.Item("mimeType"))
    If obj.Exists("lastModified") Then Let lastModified = CDbl(obj.Item("lastModified"))
    If obj.Exists("contentSize") Then Let contentSize = CDbl(obj.Item("contentSize"))
    If obj.Exists("failed") Then Let failed = CBool(obj.Item("failed"))
    If obj.Exists("canceled") Then Let canceled = CBool(obj.Item("canceled"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("url") = url
    dict("type") = fieldType
    dict("mimeType") = mimeType
    dict("lastModified") = lastModified
    dict("contentSize") = contentSize
    if Not IsEmpty(failed) Then dict("failed") = failed
    if Not IsEmpty(canceled) Then dict("canceled") = canceled

    Set asDictionary = dict
End Function

