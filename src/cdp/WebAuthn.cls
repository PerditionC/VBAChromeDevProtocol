VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAuthn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpWebAuthn"
' WebAuthn [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' 
' AuthenticatorId AS string

' 
' AuthenticatorProtocol AS string
Public Enum AuthenticatorProtocol
    AP_u2f
    AP_ctap2
End Enum

' 
' Ctap2Version AS string
Public Enum Ctap2Version
    C2V_ctap2_0
    C2V_ctap2_1
End Enum

' 
' AuthenticatorTransport AS string
Public Enum AuthenticatorTransport
    AT_usb
    AT_nfc
    AT_ble
    AT_cable
    AT_internal
End Enum

' 
' VirtualAuthenticatorOptions AS object

' 
' Credential AS object


Public Function AuthenticatorProtocolToString(ByVal ap As AuthenticatorProtocol) As String
    Dim retVal As String
    Select Case ap
        Case AP_u2f
            retVal = "u2f"
        Case AP_ctap2
            retVal = "ctap2"
        Case Else
            Debug.Print "Warning, unknown value " & ap
    End Select
    AuthenticatorProtocolToString = retVal
End Function

Public Function StringToAuthenticatorProtocol(ByVal s As String) As AuthenticatorProtocol
    Dim retVal As AuthenticatorProtocol
    Select Case s
        Case "u2f"
            retVal = AP_u2f
        Case "ctap2"
            retVal = AP_ctap2
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAuthenticatorProtocol = retVal
End Function


Public Function Ctap2VersionToString(ByVal c2v As Ctap2Version) As String
    Dim retVal As String
    Select Case c2v
        Case C2V_ctap2_0
            retVal = "ctap2_0"
        Case C2V_ctap2_1
            retVal = "ctap2_1"
        Case Else
            Debug.Print "Warning, unknown value " & c2v
    End Select
    Ctap2VersionToString = retVal
End Function

Public Function StringToCtap2Version(ByVal s As String) As Ctap2Version
    Dim retVal As Ctap2Version
    Select Case s
        Case "ctap2_0"
            retVal = C2V_ctap2_0
        Case "ctap2_1"
            retVal = C2V_ctap2_1
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCtap2Version = retVal
End Function


Public Function AuthenticatorTransportToString(ByVal at As AuthenticatorTransport) As String
    Dim retVal As String
    Select Case at
        Case AT_usb
            retVal = "usb"
        Case AT_nfc
            retVal = "nfc"
        Case AT_ble
            retVal = "ble"
        Case AT_cable
            retVal = "cable"
        Case AT_internal
            retVal = "internal"
        Case Else
            Debug.Print "Warning, unknown value " & at
    End Select
    AuthenticatorTransportToString = retVal
End Function

Public Function StringToAuthenticatorTransport(ByVal s As String) As AuthenticatorTransport
    Dim retVal As AuthenticatorTransport
    Select Case s
        Case "usb"
            retVal = AT_usb
        Case "nfc"
            retVal = AT_nfc
        Case "ble"
            retVal = AT_ble
        Case "cable"
            retVal = AT_cable
        Case "internal"
            retVal = AT_internal
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAuthenticatorTransport = retVal
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

' Enable the WebAuthn domain and start intercepting credential storage and
' retrieval with a virtual authenticator.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "WebAuthn.enable", params
End Sub

' Disable the WebAuthn domain.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "WebAuthn.disable", params
End Sub

' Creates and adds a virtual authenticator.
' Returns: authenticatorId - 
Public Function addVirtualAuthenticator( _
    ByVal options AS Dictionary _
) AS string
    ' options: VirtualAuthenticatorOptions 

    Dim params As New Dictionary
    Set params("options") = options

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("WebAuthn.addVirtualAuthenticator", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("authenticatorId") Then addVirtualAuthenticator = results("authenticatorId")
    End If
End Function

' Removes the given authenticator.
Public Sub removeVirtualAuthenticator( _
    ByVal authenticatorId AS string _
)
    ' authenticatorId: AuthenticatorId 

    Dim params As New Dictionary
    params("authenticatorId") = CStr(authenticatorId)

    cdp.InvokeMethod "WebAuthn.removeVirtualAuthenticator", params
End Sub

' Adds the credential to the specified authenticator.
Public Sub addCredential( _
    ByVal authenticatorId AS string, _
    ByVal credential AS Dictionary _
)
    ' authenticatorId: AuthenticatorId 
    ' credential: Credential 

    Dim params As New Dictionary
    params("authenticatorId") = CStr(authenticatorId)
    Set params("credential") = credential

    cdp.InvokeMethod "WebAuthn.addCredential", params
End Sub

' Returns a single credential stored in the given virtual authenticator that
' matches the credential ID.
' Returns: credential - 
Public Function getCredential( _
    ByVal authenticatorId AS string, _
    ByVal credentialId AS String _
) AS Dictionary
    ' authenticatorId: AuthenticatorId 
    ' credentialId: binary 

    Dim params As New Dictionary
    params("authenticatorId") = CStr(authenticatorId)
    params("credentialId") = CStr(credentialId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("WebAuthn.getCredential", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("credential") Then Set getCredential = results("credential")
    End If
End Function

' Returns all the credentials stored in the given virtual authenticator.
' Returns: credentials - 
Public Function getCredentials( _
    ByVal authenticatorId AS string _
) AS Collection
    ' authenticatorId: AuthenticatorId 

    Dim params As New Dictionary
    params("authenticatorId") = CStr(authenticatorId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("WebAuthn.getCredentials", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("credentials") Then Set getCredentials = results("credentials")
    End If
End Function

' Removes a credential from the authenticator.
Public Sub removeCredential( _
    ByVal authenticatorId AS string, _
    ByVal credentialId AS String _
)
    ' authenticatorId: AuthenticatorId 
    ' credentialId: binary 

    Dim params As New Dictionary
    params("authenticatorId") = CStr(authenticatorId)
    params("credentialId") = CStr(credentialId)

    cdp.InvokeMethod "WebAuthn.removeCredential", params
End Sub

' Clears all the credentials from the specified device.
Public Sub clearCredentials( _
    ByVal authenticatorId AS string _
)
    ' authenticatorId: AuthenticatorId 

    Dim params As New Dictionary
    params("authenticatorId") = CStr(authenticatorId)

    cdp.InvokeMethod "WebAuthn.clearCredentials", params
End Sub

' Sets whether User Verification succeeds or fails for an authenticator.
' The default is true.
Public Sub setUserVerified( _
    ByVal authenticatorId AS string, _
    ByVal isUserVerified AS boolean _
)
    ' authenticatorId: AuthenticatorId 
    ' isUserVerified: boolean 

    Dim params As New Dictionary
    params("authenticatorId") = CStr(authenticatorId)
    params("isUserVerified") = CBool(isUserVerified)

    cdp.InvokeMethod "WebAuthn.setUserVerified", params
End Sub

' Sets whether tests of user presence will succeed immediately (if true) or fail to resolve (if false) for an authenticator.
' The default is true.
Public Sub setAutomaticPresenceSimulation( _
    ByVal authenticatorId AS string, _
    ByVal enabled AS boolean _
)
    ' authenticatorId: AuthenticatorId 
    ' enabled: boolean 

    Dim params As New Dictionary
    params("authenticatorId") = CStr(authenticatorId)
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "WebAuthn.setAutomaticPresenceSimulation", params
End Sub

