VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpFetchRequestPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "RequestPattern"
' Fetch.RequestPattern
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' urlPattern : string
' Optional
' Wildcards (`'*'` -> zero or more, `'?'` -> exactly one) are allowed. Escape character is
    '   backslash. Omitting is equivalent to `"*"`.
Public urlPattern AS string

' resourceType : Network.ResourceType
' Optional
' If set, only requests for matching resource types will be intercepted.
Public resourceType AS string

' requestStage : RequestStage
' Optional
' Stage at which to begin intercepting requests. Default is Request.
Public requestStage AS string

Public Function init(ByVal obj as Dictionary) As cdpFetchRequestPattern
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("urlPattern") Then Let urlPattern = CStr(obj.Item("urlPattern"))
    If obj.Exists("resourceType") Then Let resourceType = CStr(obj.Item("resourceType"))
    If obj.Exists("requestStage") Then Let requestStage = CStr(obj.Item("requestStage"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if urlPattern <> vbNullString Then dict("urlPattern") = urlPattern
    if resourceType <> vbNullString Then dict("resourceType") = resourceType
    if requestStage <> vbNullString Then dict("requestStage") = requestStage

    Set asDictionary = dict
End Function

