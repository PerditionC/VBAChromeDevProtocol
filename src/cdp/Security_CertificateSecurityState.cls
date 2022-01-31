VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSecurityCertificateSecurityS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CertificateSecurityState"
' Security.CertificateSecurityState
' Details about the security state of the page certificate.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' protocol : string
' Protocol name (e.g. "TLS 1.2" or "QUIC").
Public protocol AS string

' keyExchange : string
' Key Exchange used by the connection, or the empty string if not applicable.
Public keyExchange AS string

' keyExchangeGroup : string
' Optional
' (EC)DH group used by the connection, if applicable.
Public keyExchangeGroup AS string

' cipher : string
' Cipher name.
Public cipher AS string

' mac : string
' Optional
' TLS MAC. Note that AEAD ciphers do not have separate MACs.
Public mac AS string

' certificate : array of string
' Page certificate.
Public certificate AS Collection

' subjectName : string
' Certificate subject name.
Public subjectName AS string

' issuer : string
' Name of the issuing CA.
Public issuer AS string

' validFrom : Network.TimeSinceEpoch
' Certificate valid from date.
Public validFrom AS Double

' validTo : Network.TimeSinceEpoch
' Certificate valid to (expiration) date
Public validTo AS Double

' certificateNetworkError : string
' Optional
' The highest priority network error code, if the certificate has an error.
Public certificateNetworkError AS string

' certificateHasWeakSignature : boolean
' True if the certificate uses a weak signature aglorithm.
Public certificateHasWeakSignature AS boolean

' certificateHasSha1Signature : boolean
' True if the certificate has a SHA1 signature in the chain.
Public certificateHasSha1Signature AS boolean

' modernSSL : boolean
' True if modern SSL
Public modernSSL AS boolean

' obsoleteSslProtocol : boolean
' True if the connection is using an obsolete SSL protocol.
Public obsoleteSslProtocol AS boolean

' obsoleteSslKeyExchange : boolean
' True if the connection is using an obsolete SSL key exchange.
Public obsoleteSslKeyExchange AS boolean

' obsoleteSslCipher : boolean
' True if the connection is using an obsolete SSL cipher.
Public obsoleteSslCipher AS boolean

' obsoleteSslSignature : boolean
' True if the connection is using an obsolete SSL signature.
Public obsoleteSslSignature AS boolean

Public Function init(ByVal obj as Dictionary) As cdpSecurityCertificateSecurityS
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let protocol = CStr(obj.Item("protocol"))
    Let keyExchange = CStr(obj.Item("keyExchange"))
    If obj.Exists("keyExchangeGroup") Then Let keyExchangeGroup = CStr(obj.Item("keyExchangeGroup"))
    Let cipher = CStr(obj.Item("cipher"))
    If obj.Exists("mac") Then Let mac = CStr(obj.Item("mac"))
    For each v in obj.Item("certificate")
        certificate.Add v
    Next v
    Let subjectName = CStr(obj.Item("subjectName"))
    Let issuer = CStr(obj.Item("issuer"))
    Let validFrom = CDbl(obj.Item("validFrom"))
    Let validTo = CDbl(obj.Item("validTo"))
    If obj.Exists("certificateNetworkError") Then Let certificateNetworkError = CStr(obj.Item("certificateNetworkError"))
    Let certificateHasWeakSignature = CBool(obj.Item("certificateHasWeakSignature"))
    Let certificateHasSha1Signature = CBool(obj.Item("certificateHasSha1Signature"))
    Let modernSSL = CBool(obj.Item("modernSSL"))
    Let obsoleteSslProtocol = CBool(obj.Item("obsoleteSslProtocol"))
    Let obsoleteSslKeyExchange = CBool(obj.Item("obsoleteSslKeyExchange"))
    Let obsoleteSslCipher = CBool(obj.Item("obsoleteSslCipher"))
    Let obsoleteSslSignature = CBool(obj.Item("obsoleteSslSignature"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("protocol") = protocol
    dict("keyExchange") = keyExchange
    if keyExchangeGroup <> vbNullString Then dict("keyExchangeGroup") = keyExchangeGroup
    dict("cipher") = cipher
    if mac <> vbNullString Then dict("mac") = mac
    Set dict("certificate") = certificate
    dict("subjectName") = subjectName
    dict("issuer") = issuer
    dict("validFrom") = validFrom
    dict("validTo") = validTo
    if certificateNetworkError <> vbNullString Then dict("certificateNetworkError") = certificateNetworkError
    dict("certificateHasWeakSignature") = certificateHasWeakSignature
    dict("certificateHasSha1Signature") = certificateHasSha1Signature
    dict("modernSSL") = modernSSL
    dict("obsoleteSslProtocol") = obsoleteSslProtocol
    dict("obsoleteSslKeyExchange") = obsoleteSslKeyExchange
    dict("obsoleteSslCipher") = obsoleteSslCipher
    dict("obsoleteSslSignature") = obsoleteSslSignature

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set certificate = New Collection
End Sub

