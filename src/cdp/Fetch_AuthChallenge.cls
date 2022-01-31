VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpFetchAuthChallenge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AuthChallenge"
' Fetch.AuthChallenge
' Authorization challenge for HTTP status code 401 or 407.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' source : string
' Optional
' Source of the authentication challenge.
Public source AS string

' origin : string
' Origin of the challenger.
Public origin AS string

' scheme : string
' The authentication scheme used, such as basic or digest
Public scheme AS string

' realm : string
' The realm of the challenge. May be empty.
Public realm AS string

Public Function init(ByVal obj as Dictionary) As cdpFetchAuthChallenge
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("source") Then Let source = CStr(obj.Item("source"))
    Let origin = CStr(obj.Item("origin"))
    Let scheme = CStr(obj.Item("scheme"))
    Let realm = CStr(obj.Item("realm"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if source <> vbNullString Then dict("source") = source
    dict("origin") = origin
    dict("scheme") = scheme
    dict("realm") = realm

    Set asDictionary = dict
End Function

