VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpFetchAuthChallengeResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AuthChallengeResponse"
' Fetch.AuthChallengeResponse
' Response to an AuthChallenge.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' response : string
' The decision on what to do in response to the authorization challenge.  Default means
    '   deferring to the default behavior of the net stack, which will likely either the Cancel
    '   authentication or display a popup dialog box.
Public response AS string

' username : string
' Optional
' The username to provide, possibly empty. Should only be set if response is
    '   ProvideCredentials.
Public username AS string

' password : string
' Optional
' The password to provide, possibly empty. Should only be set if response is
    '   ProvideCredentials.
Public password AS string

Public Function init(ByVal obj as Dictionary) As cdpFetchAuthChallengeResponse
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let response = CStr(obj.Item("response"))
    If obj.Exists("username") Then Let username = CStr(obj.Item("username"))
    If obj.Exists("password") Then Let password = CStr(obj.Item("password"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("response") = response
    if username <> vbNullString Then dict("username") = username
    if password <> vbNullString Then dict("password") = password

    Set asDictionary = dict
End Function

