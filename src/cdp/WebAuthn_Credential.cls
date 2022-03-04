VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAuthnCredential"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Credential"
' WebAuthn.Credential
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' credentialId : binary
' 
Public credentialId AS String

' isResidentCredential : boolean
' 
Public isResidentCredential AS boolean

' rpId : string
' Optional
' Relying Party ID the credential is scoped to. Must be set when adding a
    '   credential.
Public rpId AS string

' privateKey : binary
' The ECDSA P-256 private key in PKCS#8 format.
Public privateKey AS String

' userHandle : binary
' Optional
' An opaque byte sequence with a maximum size of 64 bytes mapping the
    '   credential to a specific user.
Public userHandle AS String

' signCount : integer
' Signature counter. This is incremented by one for each successful
    '   assertion.
    '   See https://w3c.github.io/webauthn/#signature-counter
Public signCount AS Long

' largeBlob : binary
' Optional
' The large blob associated with the credential.
    '   See https://w3c.github.io/webauthn/#sctn-large-blob-extension
Public largeBlob AS String

Public Function init(ByVal obj as Dictionary) As cdpWebAuthnCredential
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let credentialId = CStr(obj.Item("credentialId"))
    Let isResidentCredential = CBool(obj.Item("isResidentCredential"))
    If obj.Exists("rpId") Then Let rpId = CStr(obj.Item("rpId"))
    Let privateKey = CStr(obj.Item("privateKey"))
    If obj.Exists("userHandle") Then Let userHandle = CStr(obj.Item("userHandle"))
    Let signCount = CLng(obj.Item("signCount"))
    If obj.Exists("largeBlob") Then Let largeBlob = CStr(obj.Item("largeBlob"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("credentialId") = credentialId
    dict("isResidentCredential") = isResidentCredential
    if rpId <> vbNullString Then dict("rpId") = rpId
    dict("privateKey") = privateKey
    if userHandle <> vbNullString Then dict("userHandle") = userHandle
    dict("signCount") = signCount
    if largeBlob <> vbNullString Then dict("largeBlob") = largeBlob

    Set asDictionary = dict
End Function

