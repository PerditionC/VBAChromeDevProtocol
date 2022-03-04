VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpSecurity"
' Security
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' An internal certificate ID value.
' CertificateId AS integer

' A description of mixed content (HTTP resources on HTTPS pages), as defined by
'   https://www.w3.org/TR/mixed-content/#categories
' MixedContentType AS string
Public Enum MixedContentType
    MCT_blockable
    MCT_optionallyblockable
    MCT_none
End Enum

' The security level of a page or resource.
' SecurityState AS string
Public Enum SecurityState
    SS_unknown
    SS_neutral
    SS_insecure
    SS_secure
    SS_info
    SS_insecurebroken
End Enum

' Details about the security state of the page certificate.
' CertificateSecurityState AS object

' 
' SafetyTipStatus AS string
Public Enum SafetyTipStatus
    STS_badReputation
    STS_lookalike
End Enum

' 
' SafetyTipInfo AS object

' Security state information about the page.
' VisibleSecurityState AS object

' An explanation of an factor contributing to the security state.
' SecurityStateExplanation AS object

' Information about insecure content on the page.
' InsecureContentStatus AS object

' The action to take when a certificate error occurs. continue will continue processing the
'   request and cancel will cancel the request.
' CertificateErrorAction AS string
Public Enum CertificateErrorAction
    CEA_continue
    CEA_cancel
End Enum


Public Function MixedContentTypeToString(ByVal mct As MixedContentType) As String
    Dim retVal As String
    Select Case mct
        Case MCT_blockable
            retVal = "blockable"
        Case MCT_optionallyblockable
            retVal = "optionally-blockable"
        Case MCT_none
            retVal = "none"
        Case Else
            Debug.Print "Warning, unknown value " & mct
    End Select
    MixedContentTypeToString = retVal
End Function

Public Function StringToMixedContentType(ByVal s As String) As MixedContentType
    Dim retVal As MixedContentType
    Select Case s
        Case "blockable"
            retVal = MCT_blockable
        Case "optionally-blockable"
            retVal = MCT_optionallyblockable
        Case "none"
            retVal = MCT_none
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToMixedContentType = retVal
End Function


Public Function SecurityStateToString(ByVal ss As SecurityState) As String
    Dim retVal As String
    Select Case ss
        Case SS_unknown
            retVal = "unknown"
        Case SS_neutral
            retVal = "neutral"
        Case SS_insecure
            retVal = "insecure"
        Case SS_secure
            retVal = "secure"
        Case SS_info
            retVal = "info"
        Case SS_insecurebroken
            retVal = "insecure-broken"
        Case Else
            Debug.Print "Warning, unknown value " & ss
    End Select
    SecurityStateToString = retVal
End Function

Public Function StringToSecurityState(ByVal s As String) As SecurityState
    Dim retVal As SecurityState
    Select Case s
        Case "unknown"
            retVal = SS_unknown
        Case "neutral"
            retVal = SS_neutral
        Case "insecure"
            retVal = SS_insecure
        Case "secure"
            retVal = SS_secure
        Case "info"
            retVal = SS_info
        Case "insecure-broken"
            retVal = SS_insecurebroken
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSecurityState = retVal
End Function


Public Function SafetyTipStatusToString(ByVal sts As SafetyTipStatus) As String
    Dim retVal As String
    Select Case sts
        Case STS_badReputation
            retVal = "badReputation"
        Case STS_lookalike
            retVal = "lookalike"
        Case Else
            Debug.Print "Warning, unknown value " & sts
    End Select
    SafetyTipStatusToString = retVal
End Function

Public Function StringToSafetyTipStatus(ByVal s As String) As SafetyTipStatus
    Dim retVal As SafetyTipStatus
    Select Case s
        Case "badReputation"
            retVal = STS_badReputation
        Case "lookalike"
            retVal = STS_lookalike
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSafetyTipStatus = retVal
End Function


Public Function CertificateErrorActionToString(ByVal cea As CertificateErrorAction) As String
    Dim retVal As String
    Select Case cea
        Case CEA_continue
            retVal = "continue"
        Case CEA_cancel
            retVal = "cancel"
        Case Else
            Debug.Print "Warning, unknown value " & cea
    End Select
    CertificateErrorActionToString = retVal
End Function

Public Function StringToCertificateErrorAction(ByVal s As String) As CertificateErrorAction
    Dim retVal As CertificateErrorAction
    Select Case s
        Case "continue"
            retVal = CEA_continue
        Case "cancel"
            retVal = CEA_cancel
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCertificateErrorAction = retVal
End Function


Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' Disables tracking security state changes.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Security.disable", params
End Sub

' Enables tracking security state changes.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Security.enable", params
End Sub

' Enable/disable whether all certificate errors should be ignored.
' Experimental
Public Sub setIgnoreCertificateErrors( _
    ByVal ignore AS boolean _
)
    ' ignore: boolean If true, all certificate errors will be ignored.

    Dim params As New Dictionary
    params("ignore") = CBool(ignore)

    cdp.InvokeMethod "Security.setIgnoreCertificateErrors", params
End Sub

' Handles a certificate error that fired a certificateError event.
Public Sub handleCertificateError( _
    ByVal eventId AS Long, _
    ByVal action AS string _
)
    ' eventId: integer The ID of the event.
    ' action: CertificateErrorAction The action to take on the certificate error.

    Dim params As New Dictionary
    params("eventId") = CLng(eventId)
    params("action") = CStr(action)

    cdp.InvokeMethod "Security.handleCertificateError", params
End Sub

' Enable/disable overriding certificate errors. If enabled, all certificate error events need to
' be handled by the DevTools client and should be answered with `handleCertificateError` commands.
Public Sub setOverrideCertificateErrors( _
    ByVal override AS boolean _
)
    ' override: boolean If true, certificate errors will be overridden.

    Dim params As New Dictionary
    params("override") = CBool(override)

    cdp.InvokeMethod "Security.setOverrideCertificateErrors", params
End Sub

