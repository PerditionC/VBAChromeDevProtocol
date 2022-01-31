VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSecuritySecurityStateExplana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SecurityStateExplanation"
' Security.SecurityStateExplanation
' An explanation of an factor contributing to the security state.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' securityState : SecurityState
' Security state representing the severity of the factor being explained.
Public securityState AS string

' title : string
' Title describing the type of factor.
Public title AS string

' summary : string
' Short phrase describing the type of factor.
Public summary AS string

' description : string
' Full text explanation of the factor.
Public description AS string

' mixedContentType : MixedContentType
' The type of mixed content described by the explanation.
Public mixedContentType AS string

' certificate : array of string
' Page certificate.
Public certificate AS Collection

' recommendations : array of string
' Optional
' Recommendations to fix any issues.
Public recommendations AS Collection

Public Function init(ByVal obj as Dictionary) As cdpSecuritySecurityStateExplana
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let securityState = CStr(obj.Item("securityState"))
    Let title = CStr(obj.Item("title"))
    Let summary = CStr(obj.Item("summary"))
    Let description = CStr(obj.Item("description"))
    Let mixedContentType = CStr(obj.Item("mixedContentType"))
    For each v in obj.Item("certificate")
        certificate.Add v
    Next v
    If obj.Exists("recommendations") Then
        For each v in obj.Item("recommendations")
            recommendations.Add v
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("securityState") = securityState
    dict("title") = title
    dict("summary") = summary
    dict("description") = description
    dict("mixedContentType") = mixedContentType
    Set dict("certificate") = certificate
    Set dict("recommendations") = recommendations

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set certificate = New Collection
    Set recommendations = New Collection
End Sub

