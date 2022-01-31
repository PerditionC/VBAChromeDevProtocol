VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSecurityVisibleSecurityState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VisibleSecurityState"
' Security.VisibleSecurityState
' Security state information about the page.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' securityState : SecurityState
' The security level of the page.
Public securityState AS string

' certificateSecurityState : CertificateSecurityState
' Optional
' Security state details about the page certificate.
Public certificateSecurityState AS cdpSecurityCertificateSecurityS

' safetyTipInfo : SafetyTipInfo
' Optional
' The type of Safety Tip triggered on the page. Note that this field will be set even if the Safety Tip UI was not actually shown.
Public safetyTipInfo AS cdpSecuritySafetyTipInfo

' securityStateIssueIds : array of string
' Array of security state issues ids.
Public securityStateIssueIds AS Collection

Public Function init(ByVal obj as Dictionary) As cdpSecurityVisibleSecurityState
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let securityState = CStr(obj.Item("securityState"))
    If obj.Exists("certificateSecurityState") Then
        Set certificateSecurityState = New cdpSecurityCertificateSecurityS
        certificateSecurityState.init obj.Item("certificateSecurityState")
    End If
    If obj.Exists("safetyTipInfo") Then
        Set safetyTipInfo = New cdpSecuritySafetyTipInfo
        safetyTipInfo.init obj.Item("safetyTipInfo")
    End If
    For each v in obj.Item("securityStateIssueIds")
        securityStateIssueIds.Add v
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("securityState") = securityState
    If Not certificateSecurityState Is Nothing Then Set dict("certificateSecurityState") = certificateSecurityState.asDictionary()
    If Not safetyTipInfo Is Nothing Then Set dict("safetyTipInfo") = safetyTipInfo.asDictionary()
    Set dict("securityStateIssueIds") = securityStateIssueIds

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set securityStateIssueIds = New Collection
End Sub

