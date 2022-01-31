VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageOriginTrialTokenWithStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "OriginTrialTokenWithStatus"
' Page.OriginTrialTokenWithStatus
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' rawTokenText : string
' 
Public rawTokenText AS string

' parsedToken : OriginTrialToken
' Optional
' `parsedToken` is present only when the token is extractable and
    '   parsable.
Public parsedToken AS cdpPageOriginTrialToken

' status : OriginTrialTokenStatus
' 
Public status AS string

Public Function init(ByVal obj as Dictionary) As cdpPageOriginTrialTokenWithStat
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let rawTokenText = CStr(obj.Item("rawTokenText"))
    If obj.Exists("parsedToken") Then
        Set parsedToken = New cdpPageOriginTrialToken
        parsedToken.init obj.Item("parsedToken")
    End If
    Let status = CStr(obj.Item("status"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("rawTokenText") = rawTokenText
    If Not parsedToken Is Nothing Then Set dict("parsedToken") = parsedToken.asDictionary()
    dict("status") = status

    Set asDictionary = dict
End Function

