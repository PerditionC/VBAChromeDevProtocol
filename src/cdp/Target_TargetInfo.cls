VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpTargetTargetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "TargetInfo"
' Target.TargetInfo
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' targetId : TargetID
' 
Public targetId AS string

' type : string
' 
Public fieldType AS string

' title : string
' 
Public title AS string

' url : string
' 
Public url AS string

' attached : boolean
' Whether the target has an attached client.
Public attached AS boolean

' openerId : TargetID
' Optional
' Opener target Id
Public openerId AS string

' canAccessOpener : boolean
' Whether the target has access to the originating window.
Public canAccessOpener AS boolean

' openerFrameId : Page.FrameId
' Optional
' Frame id of originating window (is only set if target has an opener).
Public openerFrameId AS string

' browserContextId : Browser.BrowserContextID
' Optional
' 
Public browserContextId AS string

Public Function init(ByVal obj as Dictionary) As cdpTargetTargetInfo
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let targetId = CStr(obj.Item("targetId"))
    Let fieldType = CStr(obj.Item("type"))
    Let title = CStr(obj.Item("title"))
    Let url = CStr(obj.Item("url"))
    Let attached = CBool(obj.Item("attached"))
    If obj.Exists("openerId") Then Let openerId = CStr(obj.Item("openerId"))
    Let canAccessOpener = CBool(obj.Item("canAccessOpener"))
    If obj.Exists("openerFrameId") Then Let openerFrameId = CStr(obj.Item("openerFrameId"))
    If obj.Exists("browserContextId") Then Let browserContextId = CStr(obj.Item("browserContextId"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("targetId") = targetId
    dict("type") = fieldType
    dict("title") = title
    dict("url") = url
    dict("attached") = attached
    if openerId <> vbNullString Then dict("openerId") = openerId
    dict("canAccessOpener") = canAccessOpener
    if openerFrameId <> vbNullString Then dict("openerFrameId") = openerFrameId
    if browserContextId <> vbNullString Then dict("browserContextId") = browserContextId

    Set asDictionary = dict
End Function

