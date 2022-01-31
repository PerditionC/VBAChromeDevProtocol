VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAuthnVirtualAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VirtualAuthenticatorOptions"
' WebAuthn.VirtualAuthenticatorOptions
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' protocol : AuthenticatorProtocol
' 
Public protocol AS string

' ctap2Version : Ctap2Version
' Optional
' Defaults to ctap2_0. Ignored if |protocol| == u2f.
Public ctap2Version AS string

' transport : AuthenticatorTransport
' 
Public transport AS string

' hasResidentKey : boolean
' Optional
' Defaults to false.
Public hasResidentKey AS Variant ' boolean

' hasUserVerification : boolean
' Optional
' Defaults to false.
Public hasUserVerification AS Variant ' boolean

' hasLargeBlob : boolean
' Optional
' If set to true, the authenticator will support the largeBlob extension.
    '   https://w3c.github.io/webauthn#largeBlob
    '   Defaults to false.
Public hasLargeBlob AS Variant ' boolean

' hasCredBlob : boolean
' Optional
' If set to true, the authenticator will support the credBlob extension.
    '   https://fidoalliance.org/specs/fido-v2.1-rd-20201208/fido-client-to-authenticator-protocol-v2.1-rd-20201208.html#sctn-credBlob-extension
    '   Defaults to false.
Public hasCredBlob AS Variant ' boolean

' automaticPresenceSimulation : boolean
' Optional
' If set to true, tests of user presence will succeed immediately.
    '   Otherwise, they will not be resolved. Defaults to true.
Public automaticPresenceSimulation AS Variant ' boolean

' isUserVerified : boolean
' Optional
' Sets whether User Verification succeeds or fails for an authenticator.
    '   Defaults to false.
Public isUserVerified AS Variant ' boolean

Public Function init(ByVal obj as Dictionary) As cdpWebAuthnVirtualAuthenticator
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let protocol = CStr(obj.Item("protocol"))
    If obj.Exists("ctap2Version") Then Let ctap2Version = CStr(obj.Item("ctap2Version"))
    Let transport = CStr(obj.Item("transport"))
    If obj.Exists("hasResidentKey") Then Let hasResidentKey = CBool(obj.Item("hasResidentKey"))
    If obj.Exists("hasUserVerification") Then Let hasUserVerification = CBool(obj.Item("hasUserVerification"))
    If obj.Exists("hasLargeBlob") Then Let hasLargeBlob = CBool(obj.Item("hasLargeBlob"))
    If obj.Exists("hasCredBlob") Then Let hasCredBlob = CBool(obj.Item("hasCredBlob"))
    If obj.Exists("automaticPresenceSimulation") Then Let automaticPresenceSimulation = CBool(obj.Item("automaticPresenceSimulation"))
    If obj.Exists("isUserVerified") Then Let isUserVerified = CBool(obj.Item("isUserVerified"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("protocol") = protocol
    if ctap2Version <> vbNullString Then dict("ctap2Version") = ctap2Version
    dict("transport") = transport
    if Not IsEmpty(hasResidentKey) Then dict("hasResidentKey") = hasResidentKey
    if Not IsEmpty(hasUserVerification) Then dict("hasUserVerification") = hasUserVerification
    if Not IsEmpty(hasLargeBlob) Then dict("hasLargeBlob") = hasLargeBlob
    if Not IsEmpty(hasCredBlob) Then dict("hasCredBlob") = hasCredBlob
    if Not IsEmpty(automaticPresenceSimulation) Then dict("automaticPresenceSimulation") = automaticPresenceSimulation
    if Not IsEmpty(isUserVerified) Then dict("isUserVerified") = isUserVerified

    Set asDictionary = dict
End Function

