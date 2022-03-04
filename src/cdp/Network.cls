VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpNetwork"
' Network
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Resource type as it was perceived by the rendering engine.
' ResourceType AS string
Public Enum ResourceType
    RT_Document
    RT_Stylesheet
    RT_Image
    RT_Media
    RT_Font
    RT_Script
    RT_TextTrack
    RT_XHR
    RT_Fetch
    RT_EventSource
    RT_WebSocket
    RT_Manifest
    RT_SignedExchange
    RT_Ping
    RT_CSPViolationReport
    RT_Preflight
    RT_Other
End Enum

' Unique loader identifier.
' LoaderId AS string

' Unique request identifier.
' RequestId AS string

' Unique intercepted request identifier.
' InterceptionId AS string

' Network level fetch failure reason.
' ErrorReason AS string
Public Enum ErrorReason
    ER_Failed
    ER_Aborted
    ER_TimedOut
    ER_AccessDenied
    ER_ConnectionClosed
    ER_ConnectionReset
    ER_ConnectionRefused
    ER_ConnectionAborted
    ER_ConnectionFailed
    ER_NameNotResolved
    ER_InternetDisconnected
    ER_AddressUnreachable
    ER_BlockedByClient
    ER_BlockedByResponse
End Enum

' UTC time in seconds, counted from January 1, 1970.
' TimeSinceEpoch AS number

' Monotonically increasing time in seconds since an arbitrary point in the past.
' MonotonicTime AS number

' Request / response headers as keys / values of JSON object.
' Headers AS object

' The underlying connection technology that the browser is supposedly using.
' ConnectionType AS string
Public Enum ConnectionType
    CT_none
    CT_cellular2g
    CT_cellular3g
    CT_cellular4g
    CT_bluetooth
    CT_ethernet
    CT_wifi
    CT_wimax
    CT_other
End Enum

' Represents the cookie's 'SameSite' status:
'   https://tools.ietf.org/html/draft-west-first-party-cookies
' CookieSameSite AS string
Public Enum CookieSameSite
    CSS_Strict
    CSS_Lax
    CSS_None
End Enum

' Represents the cookie's 'Priority' status:
'   https://tools.ietf.org/html/draft-west-cookie-priority-00
' CookiePriority AS string
Public Enum CookiePriority
    CP_Low
    CP_Medium
    CP_High
End Enum

' Represents the source scheme of the origin that originally set the cookie.
'   A value of "Unset" allows protocol clients to emulate legacy cookie scope for the scheme.
'   This is a temporary ability and it will be removed in the future.
' CookieSourceScheme AS string
Public Enum CookieSourceScheme
    CSS_Unset
    CSS_NonSecure
    CSS_Secure
End Enum

' Timing information for the request.
' ResourceTiming AS object

' Loading priority of a resource request.
' ResourcePriority AS string
Public Enum ResourcePriority
    RP_VeryLow
    RP_Low
    RP_Medium
    RP_High
    RP_VeryHigh
End Enum

' Post data entry for HTTP request
' PostDataEntry AS object

' HTTP request data.
' Request AS object

' Details of a signed certificate timestamp (SCT).
' SignedCertificateTimestamp AS object

' Security details about a request.
' SecurityDetails AS object

' Whether the request complied with Certificate Transparency policy.
' CertificateTransparencyCompliance AS string
Public Enum CertificateTransparencyCompliance
    CTC_unknown
    CTC_notcompliant
    CTC_compliant
End Enum

' The reason why request was blocked.
' BlockedReason AS string
Public Enum BlockedReason
    BR_other
    BR_csp
    BR_mixedcontent
    BR_origin
    BR_inspector
    BR_subresourcefilter
    BR_contenttype
    BR_coepframeresourceneedscoepheader
    BR_coopsandboxediframecannotnavigatetocooppage
    BR_corpnotsameorigin
    BR_corpnotsameoriginafterdefaultedtosameoriginbycoep
    BR_corpnotsamesite
End Enum

' The reason why request was blocked.
' CorsError AS string
Public Enum CorsError
    CE_DisallowedByMode
    CE_InvalidResponse
    CE_WildcardOriginNotAllowed
    CE_MissingAllowOriginHeader
    CE_MultipleAllowOriginValues
    CE_InvalidAllowOriginValue
    CE_AllowOriginMismatch
    CE_InvalidAllowCredentials
    CE_CorsDisabledScheme
    CE_PreflightInvalidStatus
    CE_PreflightDisallowedRedirect
    CE_PreflightWildcardOriginNotAllowed
    CE_PreflightMissingAllowOriginHeader
    CE_PreflightMultipleAllowOriginValues
    CE_PreflightInvalidAllowOriginValue
    CE_PreflightAllowOriginMismatch
    CE_PreflightInvalidAllowCredentials
    CE_PreflightMissingAllowExternal
    CE_PreflightInvalidAllowExternal
    CE_InvalidAllowMethodsPreflightResponse
    CE_InvalidAllowHeadersPreflightResponse
    CE_MethodDisallowedByPreflightResponse
    CE_HeaderDisallowedByPreflightResponse
    CE_RedirectContainsCredentials
    CE_InsecurePrivateNetwork
    CE_InvalidPrivateNetworkAccess
    CE_UnexpectedPrivateNetworkAccess
    CE_NoCorsRedirectModeNotFollow
End Enum

' 
' CorsErrorStatus AS object

' Source of serviceworker response.
' ServiceWorkerResponseSource AS string
Public Enum ServiceWorkerResponseSource
    SWRS_cachestorage
    SWRS_httpcache
    SWRS_fallbackcode
    SWRS_network
End Enum

' Determines what type of Trust Token operation is executed and
'   depending on the type, some additional parameters. The values
'   are specified in third_party/blink/renderer/core/fetch/trust_token.idl.
' TrustTokenParams AS object

' 
' TrustTokenOperationType AS string
Public Enum TrustTokenOperationType
    TTOT_Issuance
    TTOT_Redemption
    TTOT_Signing
End Enum

' HTTP response data.
' Response AS object

' WebSocket request data.
' WebSocketRequest AS object

' WebSocket response data.
' WebSocketResponse AS object

' WebSocket message data. This represents an entire WebSocket message, not just a fragmented frame as the name suggests.
' WebSocketFrame AS object

' Information about the cached resource.
' CachedResource AS object

' Information about the request initiator.
' Initiator AS object

' Cookie object
' Cookie AS object

' Types of reasons why a cookie may not be stored from a response.
' SetCookieBlockedReason AS string
Public Enum SetCookieBlockedReason
    SCBR_SecureOnly
    SCBR_SameSiteStrict
    SCBR_SameSiteLax
    SCBR_SameSiteUnspecifiedTreatedAsLax
    SCBR_SameSiteNoneInsecure
    SCBR_UserPreferences
    SCBR_SyntaxError
    SCBR_SchemeNotSupported
    SCBR_OverwriteSecure
    SCBR_InvalidDomain
    SCBR_InvalidPrefix
    SCBR_UnknownError
    SCBR_SchemefulSameSiteStrict
    SCBR_SchemefulSameSiteLax
    SCBR_SchemefulSameSiteUnspecifiedTreatedAsLax
    SCBR_SamePartyFromCrossPartyContext
    SCBR_SamePartyConflictsWithOtherAttributes
    SCBR_NameValuePairExceedsMaxSize
End Enum

' Types of reasons why a cookie may not be sent with a request.
' CookieBlockedReason AS string
Public Enum CookieBlockedReason
    CBR_SecureOnly
    CBR_NotOnPath
    CBR_DomainMismatch
    CBR_SameSiteStrict
    CBR_SameSiteLax
    CBR_SameSiteUnspecifiedTreatedAsLax
    CBR_SameSiteNoneInsecure
    CBR_UserPreferences
    CBR_UnknownError
    CBR_SchemefulSameSiteStrict
    CBR_SchemefulSameSiteLax
    CBR_SchemefulSameSiteUnspecifiedTreatedAsLax
    CBR_SamePartyFromCrossPartyContext
    CBR_NameValuePairExceedsMaxSize
End Enum

' A cookie which was not stored from a response with the corresponding reason.
' BlockedSetCookieWithReason AS object

' A cookie with was not sent with a request with the corresponding reason.
' BlockedCookieWithReason AS object

' Cookie parameter object
' CookieParam AS object

' Authorization challenge for HTTP status code 401 or 407.
' AuthChallenge AS object

' Response to an AuthChallenge.
' AuthChallengeResponse AS object

' Stages of the interception to begin intercepting. Request will intercept before the request is
'   sent. Response will intercept after the response is received.
' InterceptionStage AS string
Public Enum InterceptionStage
    IS_Request
    IS_HeadersReceived
End Enum

' Request pattern for interception.
' RequestPattern AS object

' Information about a signed exchange signature.
'   https://wicg.github.io/webpackage/draft-yasskin-httpbis-origin-signed-exchanges-impl.html#rfc.section.3.1
' SignedExchangeSignature AS object

' Information about a signed exchange header.
'   https://wicg.github.io/webpackage/draft-yasskin-httpbis-origin-signed-exchanges-impl.html#cbor-representation
' SignedExchangeHeader AS object

' Field type for a signed exchange related error.
' SignedExchangeErrorField AS string
Public Enum SignedExchangeErrorField
    SEEF_signatureSig
    SEEF_signatureIntegrity
    SEEF_signatureCertUrl
    SEEF_signatureCertSha256
    SEEF_signatureValidityUrl
    SEEF_signatureTimestamps
End Enum

' Information about a signed exchange response.
' SignedExchangeError AS object

' Information about a signed exchange response.
' SignedExchangeInfo AS object

' List of content encodings supported by the backend.
' ContentEncoding AS string
Public Enum ContentEncoding
    CE_deflate
    CE_gzip
    CE_br
End Enum

' 
' PrivateNetworkRequestPolicy AS string
Public Enum PrivateNetworkRequestPolicy
    PNRP_Allow
    PNRP_BlockFromInsecureToMorePrivate
    PNRP_WarnFromInsecureToMorePrivate
    PNRP_PreflightBlock
    PNRP_PreflightWarn
End Enum

' 
' IPAddressSpace AS string
Public Enum IPAddressSpace
    IPAS_Local
    IPAS_Private
    IPAS_Public
    IPAS_Unknown
End Enum

' 
' ConnectTiming AS object

' 
' ClientSecurityState AS object

' 
' CrossOriginOpenerPolicyValue AS string
Public Enum CrossOriginOpenerPolicyValue
    COOPV_SameOrigin
    COOPV_SameOriginAllowPopups
    COOPV_UnsafeNone
    COOPV_SameOriginPlusCoep
End Enum

' 
' CrossOriginOpenerPolicyStatus AS object

' 
' CrossOriginEmbedderPolicyValue AS string
Public Enum CrossOriginEmbedderPolicyValue
    COEPV_None
    COEPV_Credentialless
    COEPV_RequireCorp
End Enum

' 
' CrossOriginEmbedderPolicyStatus AS object

' 
' SecurityIsolationStatus AS object

' The status of a Reporting API report.
' ReportStatus AS string
Public Enum ReportStatus
    RS_Queued
    RS_Pending
    RS_MarkedForRemoval
    RS_Success
End Enum

' 
' ReportId AS string

' An object representing a report generated by the Reporting API.
' ReportingApiReport AS object

' An object providing the result of a network resource load.
' LoadNetworkResourcePageResult AS object

' An options object that may be extended later to better support CORS,
'   CORB and streaming.
' LoadNetworkResourceOptions AS object


Public Function ResourceTypeToString(ByVal rt As ResourceType) As String
    Dim retVal As String
    Select Case rt
        Case RT_Document
            retVal = "Document"
        Case RT_Stylesheet
            retVal = "Stylesheet"
        Case RT_Image
            retVal = "Image"
        Case RT_Media
            retVal = "Media"
        Case RT_Font
            retVal = "Font"
        Case RT_Script
            retVal = "Script"
        Case RT_TextTrack
            retVal = "TextTrack"
        Case RT_XHR
            retVal = "XHR"
        Case RT_Fetch
            retVal = "Fetch"
        Case RT_EventSource
            retVal = "EventSource"
        Case RT_WebSocket
            retVal = "WebSocket"
        Case RT_Manifest
            retVal = "Manifest"
        Case RT_SignedExchange
            retVal = "SignedExchange"
        Case RT_Ping
            retVal = "Ping"
        Case RT_CSPViolationReport
            retVal = "CSPViolationReport"
        Case RT_Preflight
            retVal = "Preflight"
        Case RT_Other
            retVal = "Other"
        Case Else
            Debug.Print "Warning, unknown value " & rt
    End Select
    ResourceTypeToString = retVal
End Function

Public Function StringToResourceType(ByVal s As String) As ResourceType
    Dim retVal As ResourceType
    Select Case s
        Case "Document"
            retVal = RT_Document
        Case "Stylesheet"
            retVal = RT_Stylesheet
        Case "Image"
            retVal = RT_Image
        Case "Media"
            retVal = RT_Media
        Case "Font"
            retVal = RT_Font
        Case "Script"
            retVal = RT_Script
        Case "TextTrack"
            retVal = RT_TextTrack
        Case "XHR"
            retVal = RT_XHR
        Case "Fetch"
            retVal = RT_Fetch
        Case "EventSource"
            retVal = RT_EventSource
        Case "WebSocket"
            retVal = RT_WebSocket
        Case "Manifest"
            retVal = RT_Manifest
        Case "SignedExchange"
            retVal = RT_SignedExchange
        Case "Ping"
            retVal = RT_Ping
        Case "CSPViolationReport"
            retVal = RT_CSPViolationReport
        Case "Preflight"
            retVal = RT_Preflight
        Case "Other"
            retVal = RT_Other
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToResourceType = retVal
End Function


Public Function ErrorReasonToString(ByVal er As ErrorReason) As String
    Dim retVal As String
    Select Case er
        Case ER_Failed
            retVal = "Failed"
        Case ER_Aborted
            retVal = "Aborted"
        Case ER_TimedOut
            retVal = "TimedOut"
        Case ER_AccessDenied
            retVal = "AccessDenied"
        Case ER_ConnectionClosed
            retVal = "ConnectionClosed"
        Case ER_ConnectionReset
            retVal = "ConnectionReset"
        Case ER_ConnectionRefused
            retVal = "ConnectionRefused"
        Case ER_ConnectionAborted
            retVal = "ConnectionAborted"
        Case ER_ConnectionFailed
            retVal = "ConnectionFailed"
        Case ER_NameNotResolved
            retVal = "NameNotResolved"
        Case ER_InternetDisconnected
            retVal = "InternetDisconnected"
        Case ER_AddressUnreachable
            retVal = "AddressUnreachable"
        Case ER_BlockedByClient
            retVal = "BlockedByClient"
        Case ER_BlockedByResponse
            retVal = "BlockedByResponse"
        Case Else
            Debug.Print "Warning, unknown value " & er
    End Select
    ErrorReasonToString = retVal
End Function

Public Function StringToErrorReason(ByVal s As String) As ErrorReason
    Dim retVal As ErrorReason
    Select Case s
        Case "Failed"
            retVal = ER_Failed
        Case "Aborted"
            retVal = ER_Aborted
        Case "TimedOut"
            retVal = ER_TimedOut
        Case "AccessDenied"
            retVal = ER_AccessDenied
        Case "ConnectionClosed"
            retVal = ER_ConnectionClosed
        Case "ConnectionReset"
            retVal = ER_ConnectionReset
        Case "ConnectionRefused"
            retVal = ER_ConnectionRefused
        Case "ConnectionAborted"
            retVal = ER_ConnectionAborted
        Case "ConnectionFailed"
            retVal = ER_ConnectionFailed
        Case "NameNotResolved"
            retVal = ER_NameNotResolved
        Case "InternetDisconnected"
            retVal = ER_InternetDisconnected
        Case "AddressUnreachable"
            retVal = ER_AddressUnreachable
        Case "BlockedByClient"
            retVal = ER_BlockedByClient
        Case "BlockedByResponse"
            retVal = ER_BlockedByResponse
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToErrorReason = retVal
End Function


Public Function ConnectionTypeToString(ByVal ct As ConnectionType) As String
    Dim retVal As String
    Select Case ct
        Case CT_none
            retVal = "none"
        Case CT_cellular2g
            retVal = "cellular2g"
        Case CT_cellular3g
            retVal = "cellular3g"
        Case CT_cellular4g
            retVal = "cellular4g"
        Case CT_bluetooth
            retVal = "bluetooth"
        Case CT_ethernet
            retVal = "ethernet"
        Case CT_wifi
            retVal = "wifi"
        Case CT_wimax
            retVal = "wimax"
        Case CT_other
            retVal = "other"
        Case Else
            Debug.Print "Warning, unknown value " & ct
    End Select
    ConnectionTypeToString = retVal
End Function

Public Function StringToConnectionType(ByVal s As String) As ConnectionType
    Dim retVal As ConnectionType
    Select Case s
        Case "none"
            retVal = CT_none
        Case "cellular2g"
            retVal = CT_cellular2g
        Case "cellular3g"
            retVal = CT_cellular3g
        Case "cellular4g"
            retVal = CT_cellular4g
        Case "bluetooth"
            retVal = CT_bluetooth
        Case "ethernet"
            retVal = CT_ethernet
        Case "wifi"
            retVal = CT_wifi
        Case "wimax"
            retVal = CT_wimax
        Case "other"
            retVal = CT_other
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToConnectionType = retVal
End Function


Public Function CookieSameSiteToString(ByVal css As CookieSameSite) As String
    Dim retVal As String
    Select Case css
        Case CSS_Strict
            retVal = "Strict"
        Case CSS_Lax
            retVal = "Lax"
        Case CSS_None
            retVal = "None"
        Case Else
            Debug.Print "Warning, unknown value " & css
    End Select
    CookieSameSiteToString = retVal
End Function

Public Function StringToCookieSameSite(ByVal s As String) As CookieSameSite
    Dim retVal As CookieSameSite
    Select Case s
        Case "Strict"
            retVal = CSS_Strict
        Case "Lax"
            retVal = CSS_Lax
        Case "None"
            retVal = CSS_None
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCookieSameSite = retVal
End Function


Public Function CookiePriorityToString(ByVal cp As CookiePriority) As String
    Dim retVal As String
    Select Case cp
        Case CP_Low
            retVal = "Low"
        Case CP_Medium
            retVal = "Medium"
        Case CP_High
            retVal = "High"
        Case Else
            Debug.Print "Warning, unknown value " & cp
    End Select
    CookiePriorityToString = retVal
End Function

Public Function StringToCookiePriority(ByVal s As String) As CookiePriority
    Dim retVal As CookiePriority
    Select Case s
        Case "Low"
            retVal = CP_Low
        Case "Medium"
            retVal = CP_Medium
        Case "High"
            retVal = CP_High
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCookiePriority = retVal
End Function


Public Function CookieSourceSchemeToString(ByVal css As CookieSourceScheme) As String
    Dim retVal As String
    Select Case css
        Case CSS_Unset
            retVal = "Unset"
        Case CSS_NonSecure
            retVal = "NonSecure"
        Case CSS_Secure
            retVal = "Secure"
        Case Else
            Debug.Print "Warning, unknown value " & css
    End Select
    CookieSourceSchemeToString = retVal
End Function

Public Function StringToCookieSourceScheme(ByVal s As String) As CookieSourceScheme
    Dim retVal As CookieSourceScheme
    Select Case s
        Case "Unset"
            retVal = CSS_Unset
        Case "NonSecure"
            retVal = CSS_NonSecure
        Case "Secure"
            retVal = CSS_Secure
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCookieSourceScheme = retVal
End Function


Public Function ResourcePriorityToString(ByVal rp As ResourcePriority) As String
    Dim retVal As String
    Select Case rp
        Case RP_VeryLow
            retVal = "VeryLow"
        Case RP_Low
            retVal = "Low"
        Case RP_Medium
            retVal = "Medium"
        Case RP_High
            retVal = "High"
        Case RP_VeryHigh
            retVal = "VeryHigh"
        Case Else
            Debug.Print "Warning, unknown value " & rp
    End Select
    ResourcePriorityToString = retVal
End Function

Public Function StringToResourcePriority(ByVal s As String) As ResourcePriority
    Dim retVal As ResourcePriority
    Select Case s
        Case "VeryLow"
            retVal = RP_VeryLow
        Case "Low"
            retVal = RP_Low
        Case "Medium"
            retVal = RP_Medium
        Case "High"
            retVal = RP_High
        Case "VeryHigh"
            retVal = RP_VeryHigh
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToResourcePriority = retVal
End Function


Public Function CertificateTransparencyComplianceToString(ByVal ctc As CertificateTransparencyCompliance) As String
    Dim retVal As String
    Select Case ctc
        Case CTC_unknown
            retVal = "unknown"
        Case CTC_notcompliant
            retVal = "not-compliant"
        Case CTC_compliant
            retVal = "compliant"
        Case Else
            Debug.Print "Warning, unknown value " & ctc
    End Select
    CertificateTransparencyComplianceToString = retVal
End Function

Public Function StringToCertificateTransparencyCompliance(ByVal s As String) As CertificateTransparencyCompliance
    Dim retVal As CertificateTransparencyCompliance
    Select Case s
        Case "unknown"
            retVal = CTC_unknown
        Case "not-compliant"
            retVal = CTC_notcompliant
        Case "compliant"
            retVal = CTC_compliant
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCertificateTransparencyCompliance = retVal
End Function


Public Function BlockedReasonToString(ByVal br As BlockedReason) As String
    Dim retVal As String
    Select Case br
        Case BR_other
            retVal = "other"
        Case BR_csp
            retVal = "csp"
        Case BR_mixedcontent
            retVal = "mixed-content"
        Case BR_origin
            retVal = "origin"
        Case BR_inspector
            retVal = "inspector"
        Case BR_subresourcefilter
            retVal = "subresource-filter"
        Case BR_contenttype
            retVal = "content-type"
        Case BR_coepframeresourceneedscoepheader
            retVal = "coep-frame-resource-needs-coep-header"
        Case BR_coopsandboxediframecannotnavigatetocooppage
            retVal = "coop-sandboxed-iframe-cannot-navigate-to-coop-page"
        Case BR_corpnotsameorigin
            retVal = "corp-not-same-origin"
        Case BR_corpnotsameoriginafterdefaultedtosameoriginbycoep
            retVal = "corp-not-same-origin-after-defaulted-to-same-origin-by-coep"
        Case BR_corpnotsamesite
            retVal = "corp-not-same-site"
        Case Else
            Debug.Print "Warning, unknown value " & br
    End Select
    BlockedReasonToString = retVal
End Function

Public Function StringToBlockedReason(ByVal s As String) As BlockedReason
    Dim retVal As BlockedReason
    Select Case s
        Case "other"
            retVal = BR_other
        Case "csp"
            retVal = BR_csp
        Case "mixed-content"
            retVal = BR_mixedcontent
        Case "origin"
            retVal = BR_origin
        Case "inspector"
            retVal = BR_inspector
        Case "subresource-filter"
            retVal = BR_subresourcefilter
        Case "content-type"
            retVal = BR_contenttype
        Case "coep-frame-resource-needs-coep-header"
            retVal = BR_coepframeresourceneedscoepheader
        Case "coop-sandboxed-iframe-cannot-navigate-to-coop-page"
            retVal = BR_coopsandboxediframecannotnavigatetocooppage
        Case "corp-not-same-origin"
            retVal = BR_corpnotsameorigin
        Case "corp-not-same-origin-after-defaulted-to-same-origin-by-coep"
            retVal = BR_corpnotsameoriginafterdefaultedtosameoriginbycoep
        Case "corp-not-same-site"
            retVal = BR_corpnotsamesite
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToBlockedReason = retVal
End Function


Public Function CorsErrorToString(ByVal ce As CorsError) As String
    Dim retVal As String
    Select Case ce
        Case CE_DisallowedByMode
            retVal = "DisallowedByMode"
        Case CE_InvalidResponse
            retVal = "InvalidResponse"
        Case CE_WildcardOriginNotAllowed
            retVal = "WildcardOriginNotAllowed"
        Case CE_MissingAllowOriginHeader
            retVal = "MissingAllowOriginHeader"
        Case CE_MultipleAllowOriginValues
            retVal = "MultipleAllowOriginValues"
        Case CE_InvalidAllowOriginValue
            retVal = "InvalidAllowOriginValue"
        Case CE_AllowOriginMismatch
            retVal = "AllowOriginMismatch"
        Case CE_InvalidAllowCredentials
            retVal = "InvalidAllowCredentials"
        Case CE_CorsDisabledScheme
            retVal = "CorsDisabledScheme"
        Case CE_PreflightInvalidStatus
            retVal = "PreflightInvalidStatus"
        Case CE_PreflightDisallowedRedirect
            retVal = "PreflightDisallowedRedirect"
        Case CE_PreflightWildcardOriginNotAllowed
            retVal = "PreflightWildcardOriginNotAllowed"
        Case CE_PreflightMissingAllowOriginHeader
            retVal = "PreflightMissingAllowOriginHeader"
        Case CE_PreflightMultipleAllowOriginValues
            retVal = "PreflightMultipleAllowOriginValues"
        Case CE_PreflightInvalidAllowOriginValue
            retVal = "PreflightInvalidAllowOriginValue"
        Case CE_PreflightAllowOriginMismatch
            retVal = "PreflightAllowOriginMismatch"
        Case CE_PreflightInvalidAllowCredentials
            retVal = "PreflightInvalidAllowCredentials"
        Case CE_PreflightMissingAllowExternal
            retVal = "PreflightMissingAllowExternal"
        Case CE_PreflightInvalidAllowExternal
            retVal = "PreflightInvalidAllowExternal"
        Case CE_InvalidAllowMethodsPreflightResponse
            retVal = "InvalidAllowMethodsPreflightResponse"
        Case CE_InvalidAllowHeadersPreflightResponse
            retVal = "InvalidAllowHeadersPreflightResponse"
        Case CE_MethodDisallowedByPreflightResponse
            retVal = "MethodDisallowedByPreflightResponse"
        Case CE_HeaderDisallowedByPreflightResponse
            retVal = "HeaderDisallowedByPreflightResponse"
        Case CE_RedirectContainsCredentials
            retVal = "RedirectContainsCredentials"
        Case CE_InsecurePrivateNetwork
            retVal = "InsecurePrivateNetwork"
        Case CE_InvalidPrivateNetworkAccess
            retVal = "InvalidPrivateNetworkAccess"
        Case CE_UnexpectedPrivateNetworkAccess
            retVal = "UnexpectedPrivateNetworkAccess"
        Case CE_NoCorsRedirectModeNotFollow
            retVal = "NoCorsRedirectModeNotFollow"
        Case Else
            Debug.Print "Warning, unknown value " & ce
    End Select
    CorsErrorToString = retVal
End Function

Public Function StringToCorsError(ByVal s As String) As CorsError
    Dim retVal As CorsError
    Select Case s
        Case "DisallowedByMode"
            retVal = CE_DisallowedByMode
        Case "InvalidResponse"
            retVal = CE_InvalidResponse
        Case "WildcardOriginNotAllowed"
            retVal = CE_WildcardOriginNotAllowed
        Case "MissingAllowOriginHeader"
            retVal = CE_MissingAllowOriginHeader
        Case "MultipleAllowOriginValues"
            retVal = CE_MultipleAllowOriginValues
        Case "InvalidAllowOriginValue"
            retVal = CE_InvalidAllowOriginValue
        Case "AllowOriginMismatch"
            retVal = CE_AllowOriginMismatch
        Case "InvalidAllowCredentials"
            retVal = CE_InvalidAllowCredentials
        Case "CorsDisabledScheme"
            retVal = CE_CorsDisabledScheme
        Case "PreflightInvalidStatus"
            retVal = CE_PreflightInvalidStatus
        Case "PreflightDisallowedRedirect"
            retVal = CE_PreflightDisallowedRedirect
        Case "PreflightWildcardOriginNotAllowed"
            retVal = CE_PreflightWildcardOriginNotAllowed
        Case "PreflightMissingAllowOriginHeader"
            retVal = CE_PreflightMissingAllowOriginHeader
        Case "PreflightMultipleAllowOriginValues"
            retVal = CE_PreflightMultipleAllowOriginValues
        Case "PreflightInvalidAllowOriginValue"
            retVal = CE_PreflightInvalidAllowOriginValue
        Case "PreflightAllowOriginMismatch"
            retVal = CE_PreflightAllowOriginMismatch
        Case "PreflightInvalidAllowCredentials"
            retVal = CE_PreflightInvalidAllowCredentials
        Case "PreflightMissingAllowExternal"
            retVal = CE_PreflightMissingAllowExternal
        Case "PreflightInvalidAllowExternal"
            retVal = CE_PreflightInvalidAllowExternal
        Case "InvalidAllowMethodsPreflightResponse"
            retVal = CE_InvalidAllowMethodsPreflightResponse
        Case "InvalidAllowHeadersPreflightResponse"
            retVal = CE_InvalidAllowHeadersPreflightResponse
        Case "MethodDisallowedByPreflightResponse"
            retVal = CE_MethodDisallowedByPreflightResponse
        Case "HeaderDisallowedByPreflightResponse"
            retVal = CE_HeaderDisallowedByPreflightResponse
        Case "RedirectContainsCredentials"
            retVal = CE_RedirectContainsCredentials
        Case "InsecurePrivateNetwork"
            retVal = CE_InsecurePrivateNetwork
        Case "InvalidPrivateNetworkAccess"
            retVal = CE_InvalidPrivateNetworkAccess
        Case "UnexpectedPrivateNetworkAccess"
            retVal = CE_UnexpectedPrivateNetworkAccess
        Case "NoCorsRedirectModeNotFollow"
            retVal = CE_NoCorsRedirectModeNotFollow
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCorsError = retVal
End Function


Public Function ServiceWorkerResponseSourceToString(ByVal swrs As ServiceWorkerResponseSource) As String
    Dim retVal As String
    Select Case swrs
        Case SWRS_cachestorage
            retVal = "cache-storage"
        Case SWRS_httpcache
            retVal = "http-cache"
        Case SWRS_fallbackcode
            retVal = "fallback-code"
        Case SWRS_network
            retVal = "network"
        Case Else
            Debug.Print "Warning, unknown value " & swrs
    End Select
    ServiceWorkerResponseSourceToString = retVal
End Function

Public Function StringToServiceWorkerResponseSource(ByVal s As String) As ServiceWorkerResponseSource
    Dim retVal As ServiceWorkerResponseSource
    Select Case s
        Case "cache-storage"
            retVal = SWRS_cachestorage
        Case "http-cache"
            retVal = SWRS_httpcache
        Case "fallback-code"
            retVal = SWRS_fallbackcode
        Case "network"
            retVal = SWRS_network
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToServiceWorkerResponseSource = retVal
End Function


Public Function TrustTokenOperationTypeToString(ByVal ttot As TrustTokenOperationType) As String
    Dim retVal As String
    Select Case ttot
        Case TTOT_Issuance
            retVal = "Issuance"
        Case TTOT_Redemption
            retVal = "Redemption"
        Case TTOT_Signing
            retVal = "Signing"
        Case Else
            Debug.Print "Warning, unknown value " & ttot
    End Select
    TrustTokenOperationTypeToString = retVal
End Function

Public Function StringToTrustTokenOperationType(ByVal s As String) As TrustTokenOperationType
    Dim retVal As TrustTokenOperationType
    Select Case s
        Case "Issuance"
            retVal = TTOT_Issuance
        Case "Redemption"
            retVal = TTOT_Redemption
        Case "Signing"
            retVal = TTOT_Signing
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToTrustTokenOperationType = retVal
End Function


Public Function SetCookieBlockedReasonToString(ByVal scbr As SetCookieBlockedReason) As String
    Dim retVal As String
    Select Case scbr
        Case SCBR_SecureOnly
            retVal = "SecureOnly"
        Case SCBR_SameSiteStrict
            retVal = "SameSiteStrict"
        Case SCBR_SameSiteLax
            retVal = "SameSiteLax"
        Case SCBR_SameSiteUnspecifiedTreatedAsLax
            retVal = "SameSiteUnspecifiedTreatedAsLax"
        Case SCBR_SameSiteNoneInsecure
            retVal = "SameSiteNoneInsecure"
        Case SCBR_UserPreferences
            retVal = "UserPreferences"
        Case SCBR_SyntaxError
            retVal = "SyntaxError"
        Case SCBR_SchemeNotSupported
            retVal = "SchemeNotSupported"
        Case SCBR_OverwriteSecure
            retVal = "OverwriteSecure"
        Case SCBR_InvalidDomain
            retVal = "InvalidDomain"
        Case SCBR_InvalidPrefix
            retVal = "InvalidPrefix"
        Case SCBR_UnknownError
            retVal = "UnknownError"
        Case SCBR_SchemefulSameSiteStrict
            retVal = "SchemefulSameSiteStrict"
        Case SCBR_SchemefulSameSiteLax
            retVal = "SchemefulSameSiteLax"
        Case SCBR_SchemefulSameSiteUnspecifiedTreatedAsLax
            retVal = "SchemefulSameSiteUnspecifiedTreatedAsLax"
        Case SCBR_SamePartyFromCrossPartyContext
            retVal = "SamePartyFromCrossPartyContext"
        Case SCBR_SamePartyConflictsWithOtherAttributes
            retVal = "SamePartyConflictsWithOtherAttributes"
        Case SCBR_NameValuePairExceedsMaxSize
            retVal = "NameValuePairExceedsMaxSize"
        Case Else
            Debug.Print "Warning, unknown value " & scbr
    End Select
    SetCookieBlockedReasonToString = retVal
End Function

Public Function StringToSetCookieBlockedReason(ByVal s As String) As SetCookieBlockedReason
    Dim retVal As SetCookieBlockedReason
    Select Case s
        Case "SecureOnly"
            retVal = SCBR_SecureOnly
        Case "SameSiteStrict"
            retVal = SCBR_SameSiteStrict
        Case "SameSiteLax"
            retVal = SCBR_SameSiteLax
        Case "SameSiteUnspecifiedTreatedAsLax"
            retVal = SCBR_SameSiteUnspecifiedTreatedAsLax
        Case "SameSiteNoneInsecure"
            retVal = SCBR_SameSiteNoneInsecure
        Case "UserPreferences"
            retVal = SCBR_UserPreferences
        Case "SyntaxError"
            retVal = SCBR_SyntaxError
        Case "SchemeNotSupported"
            retVal = SCBR_SchemeNotSupported
        Case "OverwriteSecure"
            retVal = SCBR_OverwriteSecure
        Case "InvalidDomain"
            retVal = SCBR_InvalidDomain
        Case "InvalidPrefix"
            retVal = SCBR_InvalidPrefix
        Case "UnknownError"
            retVal = SCBR_UnknownError
        Case "SchemefulSameSiteStrict"
            retVal = SCBR_SchemefulSameSiteStrict
        Case "SchemefulSameSiteLax"
            retVal = SCBR_SchemefulSameSiteLax
        Case "SchemefulSameSiteUnspecifiedTreatedAsLax"
            retVal = SCBR_SchemefulSameSiteUnspecifiedTreatedAsLax
        Case "SamePartyFromCrossPartyContext"
            retVal = SCBR_SamePartyFromCrossPartyContext
        Case "SamePartyConflictsWithOtherAttributes"
            retVal = SCBR_SamePartyConflictsWithOtherAttributes
        Case "NameValuePairExceedsMaxSize"
            retVal = SCBR_NameValuePairExceedsMaxSize
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSetCookieBlockedReason = retVal
End Function


Public Function CookieBlockedReasonToString(ByVal cbr As CookieBlockedReason) As String
    Dim retVal As String
    Select Case cbr
        Case CBR_SecureOnly
            retVal = "SecureOnly"
        Case CBR_NotOnPath
            retVal = "NotOnPath"
        Case CBR_DomainMismatch
            retVal = "DomainMismatch"
        Case CBR_SameSiteStrict
            retVal = "SameSiteStrict"
        Case CBR_SameSiteLax
            retVal = "SameSiteLax"
        Case CBR_SameSiteUnspecifiedTreatedAsLax
            retVal = "SameSiteUnspecifiedTreatedAsLax"
        Case CBR_SameSiteNoneInsecure
            retVal = "SameSiteNoneInsecure"
        Case CBR_UserPreferences
            retVal = "UserPreferences"
        Case CBR_UnknownError
            retVal = "UnknownError"
        Case CBR_SchemefulSameSiteStrict
            retVal = "SchemefulSameSiteStrict"
        Case CBR_SchemefulSameSiteLax
            retVal = "SchemefulSameSiteLax"
        Case CBR_SchemefulSameSiteUnspecifiedTreatedAsLax
            retVal = "SchemefulSameSiteUnspecifiedTreatedAsLax"
        Case CBR_SamePartyFromCrossPartyContext
            retVal = "SamePartyFromCrossPartyContext"
        Case CBR_NameValuePairExceedsMaxSize
            retVal = "NameValuePairExceedsMaxSize"
        Case Else
            Debug.Print "Warning, unknown value " & cbr
    End Select
    CookieBlockedReasonToString = retVal
End Function

Public Function StringToCookieBlockedReason(ByVal s As String) As CookieBlockedReason
    Dim retVal As CookieBlockedReason
    Select Case s
        Case "SecureOnly"
            retVal = CBR_SecureOnly
        Case "NotOnPath"
            retVal = CBR_NotOnPath
        Case "DomainMismatch"
            retVal = CBR_DomainMismatch
        Case "SameSiteStrict"
            retVal = CBR_SameSiteStrict
        Case "SameSiteLax"
            retVal = CBR_SameSiteLax
        Case "SameSiteUnspecifiedTreatedAsLax"
            retVal = CBR_SameSiteUnspecifiedTreatedAsLax
        Case "SameSiteNoneInsecure"
            retVal = CBR_SameSiteNoneInsecure
        Case "UserPreferences"
            retVal = CBR_UserPreferences
        Case "UnknownError"
            retVal = CBR_UnknownError
        Case "SchemefulSameSiteStrict"
            retVal = CBR_SchemefulSameSiteStrict
        Case "SchemefulSameSiteLax"
            retVal = CBR_SchemefulSameSiteLax
        Case "SchemefulSameSiteUnspecifiedTreatedAsLax"
            retVal = CBR_SchemefulSameSiteUnspecifiedTreatedAsLax
        Case "SamePartyFromCrossPartyContext"
            retVal = CBR_SamePartyFromCrossPartyContext
        Case "NameValuePairExceedsMaxSize"
            retVal = CBR_NameValuePairExceedsMaxSize
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCookieBlockedReason = retVal
End Function


Public Function InterceptionStageToString(ByVal e As InterceptionStage) As String
    Dim retVal As String
    Select Case e
        Case IS_Request
            retVal = "Request"
        Case IS_HeadersReceived
            retVal = "HeadersReceived"
        Case Else
            Debug.Print "Warning, unknown value " & e
    End Select
    InterceptionStageToString = retVal
End Function

Public Function StringToInterceptionStage(ByVal s As String) As InterceptionStage
    Dim retVal As InterceptionStage
    Select Case s
        Case "Request"
            retVal = IS_Request
        Case "HeadersReceived"
            retVal = IS_HeadersReceived
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToInterceptionStage = retVal
End Function


Public Function SignedExchangeErrorFieldToString(ByVal seef As SignedExchangeErrorField) As String
    Dim retVal As String
    Select Case seef
        Case SEEF_signatureSig
            retVal = "signatureSig"
        Case SEEF_signatureIntegrity
            retVal = "signatureIntegrity"
        Case SEEF_signatureCertUrl
            retVal = "signatureCertUrl"
        Case SEEF_signatureCertSha256
            retVal = "signatureCertSha256"
        Case SEEF_signatureValidityUrl
            retVal = "signatureValidityUrl"
        Case SEEF_signatureTimestamps
            retVal = "signatureTimestamps"
        Case Else
            Debug.Print "Warning, unknown value " & seef
    End Select
    SignedExchangeErrorFieldToString = retVal
End Function

Public Function StringToSignedExchangeErrorField(ByVal s As String) As SignedExchangeErrorField
    Dim retVal As SignedExchangeErrorField
    Select Case s
        Case "signatureSig"
            retVal = SEEF_signatureSig
        Case "signatureIntegrity"
            retVal = SEEF_signatureIntegrity
        Case "signatureCertUrl"
            retVal = SEEF_signatureCertUrl
        Case "signatureCertSha256"
            retVal = SEEF_signatureCertSha256
        Case "signatureValidityUrl"
            retVal = SEEF_signatureValidityUrl
        Case "signatureTimestamps"
            retVal = SEEF_signatureTimestamps
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSignedExchangeErrorField = retVal
End Function


Public Function ContentEncodingToString(ByVal ce As ContentEncoding) As String
    Dim retVal As String
    Select Case ce
        Case CE_deflate
            retVal = "deflate"
        Case CE_gzip
            retVal = "gzip"
        Case CE_br
            retVal = "br"
        Case Else
            Debug.Print "Warning, unknown value " & ce
    End Select
    ContentEncodingToString = retVal
End Function

Public Function StringToContentEncoding(ByVal s As String) As ContentEncoding
    Dim retVal As ContentEncoding
    Select Case s
        Case "deflate"
            retVal = CE_deflate
        Case "gzip"
            retVal = CE_gzip
        Case "br"
            retVal = CE_br
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToContentEncoding = retVal
End Function


Public Function PrivateNetworkRequestPolicyToString(ByVal pnrp As PrivateNetworkRequestPolicy) As String
    Dim retVal As String
    Select Case pnrp
        Case PNRP_Allow
            retVal = "Allow"
        Case PNRP_BlockFromInsecureToMorePrivate
            retVal = "BlockFromInsecureToMorePrivate"
        Case PNRP_WarnFromInsecureToMorePrivate
            retVal = "WarnFromInsecureToMorePrivate"
        Case PNRP_PreflightBlock
            retVal = "PreflightBlock"
        Case PNRP_PreflightWarn
            retVal = "PreflightWarn"
        Case Else
            Debug.Print "Warning, unknown value " & pnrp
    End Select
    PrivateNetworkRequestPolicyToString = retVal
End Function

Public Function StringToPrivateNetworkRequestPolicy(ByVal s As String) As PrivateNetworkRequestPolicy
    Dim retVal As PrivateNetworkRequestPolicy
    Select Case s
        Case "Allow"
            retVal = PNRP_Allow
        Case "BlockFromInsecureToMorePrivate"
            retVal = PNRP_BlockFromInsecureToMorePrivate
        Case "WarnFromInsecureToMorePrivate"
            retVal = PNRP_WarnFromInsecureToMorePrivate
        Case "PreflightBlock"
            retVal = PNRP_PreflightBlock
        Case "PreflightWarn"
            retVal = PNRP_PreflightWarn
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToPrivateNetworkRequestPolicy = retVal
End Function


Public Function IPAddressSpaceToString(ByVal ipas As IPAddressSpace) As String
    Dim retVal As String
    Select Case ipas
        Case IPAS_Local
            retVal = "Local"
        Case IPAS_Private
            retVal = "Private"
        Case IPAS_Public
            retVal = "Public"
        Case IPAS_Unknown
            retVal = "Unknown"
        Case Else
            Debug.Print "Warning, unknown value " & ipas
    End Select
    IPAddressSpaceToString = retVal
End Function

Public Function StringToIPAddressSpace(ByVal s As String) As IPAddressSpace
    Dim retVal As IPAddressSpace
    Select Case s
        Case "Local"
            retVal = IPAS_Local
        Case "Private"
            retVal = IPAS_Private
        Case "Public"
            retVal = IPAS_Public
        Case "Unknown"
            retVal = IPAS_Unknown
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToIPAddressSpace = retVal
End Function


Public Function CrossOriginOpenerPolicyValueToString(ByVal coopv As CrossOriginOpenerPolicyValue) As String
    Dim retVal As String
    Select Case coopv
        Case COOPV_SameOrigin
            retVal = "SameOrigin"
        Case COOPV_SameOriginAllowPopups
            retVal = "SameOriginAllowPopups"
        Case COOPV_UnsafeNone
            retVal = "UnsafeNone"
        Case COOPV_SameOriginPlusCoep
            retVal = "SameOriginPlusCoep"
        Case Else
            Debug.Print "Warning, unknown value " & coopv
    End Select
    CrossOriginOpenerPolicyValueToString = retVal
End Function

Public Function StringToCrossOriginOpenerPolicyValue(ByVal s As String) As CrossOriginOpenerPolicyValue
    Dim retVal As CrossOriginOpenerPolicyValue
    Select Case s
        Case "SameOrigin"
            retVal = COOPV_SameOrigin
        Case "SameOriginAllowPopups"
            retVal = COOPV_SameOriginAllowPopups
        Case "UnsafeNone"
            retVal = COOPV_UnsafeNone
        Case "SameOriginPlusCoep"
            retVal = COOPV_SameOriginPlusCoep
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCrossOriginOpenerPolicyValue = retVal
End Function


Public Function CrossOriginEmbedderPolicyValueToString(ByVal coepv As CrossOriginEmbedderPolicyValue) As String
    Dim retVal As String
    Select Case coepv
        Case COEPV_None
            retVal = "None"
        Case COEPV_Credentialless
            retVal = "Credentialless"
        Case COEPV_RequireCorp
            retVal = "RequireCorp"
        Case Else
            Debug.Print "Warning, unknown value " & coepv
    End Select
    CrossOriginEmbedderPolicyValueToString = retVal
End Function

Public Function StringToCrossOriginEmbedderPolicyValue(ByVal s As String) As CrossOriginEmbedderPolicyValue
    Dim retVal As CrossOriginEmbedderPolicyValue
    Select Case s
        Case "None"
            retVal = COEPV_None
        Case "Credentialless"
            retVal = COEPV_Credentialless
        Case "RequireCorp"
            retVal = COEPV_RequireCorp
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCrossOriginEmbedderPolicyValue = retVal
End Function


Public Function ReportStatusToString(ByVal rs As ReportStatus) As String
    Dim retVal As String
    Select Case rs
        Case RS_Queued
            retVal = "Queued"
        Case RS_Pending
            retVal = "Pending"
        Case RS_MarkedForRemoval
            retVal = "MarkedForRemoval"
        Case RS_Success
            retVal = "Success"
        Case Else
            Debug.Print "Warning, unknown value " & rs
    End Select
    ReportStatusToString = retVal
End Function

Public Function StringToReportStatus(ByVal s As String) As ReportStatus
    Dim retVal As ReportStatus
    Select Case s
        Case "Queued"
            retVal = RS_Queued
        Case "Pending"
            retVal = RS_Pending
        Case "MarkedForRemoval"
            retVal = RS_MarkedForRemoval
        Case "Success"
            retVal = RS_Success
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToReportStatus = retVal
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

' Sets a list of content encodings that will be accepted. Empty list means no encoding is accepted.
' Experimental
Public Sub setAcceptedEncodings( _
    ByVal encodings AS Collection _
)
    ' encodings: array List of accepted content encodings.

    Dim params As New Dictionary
    Set params("encodings") = encodings

    cdp.InvokeMethod "Network.setAcceptedEncodings", params
End Sub

' Clears accepted encodings set by setAcceptedEncodings
' Experimental
Public Sub clearAcceptedEncodingsOverride()
    Dim params As New Dictionary
    cdp.InvokeMethod "Network.clearAcceptedEncodingsOverride", params
End Sub

' Tells whether clearing browser cache is supported.
' Returns: result - True if browser cache can be cleared.
Public Function canClearBrowserCache() AS boolean
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.canClearBrowserCache", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then canClearBrowserCache = results("result")
    End If
End Function

' Tells whether clearing browser cookies is supported.
' Returns: result - True if browser cookies can be cleared.
Public Function canClearBrowserCookies() AS boolean
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.canClearBrowserCookies", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then canClearBrowserCookies = results("result")
    End If
End Function

' Tells whether emulation of network conditions is supported.
' Returns: result - True if emulation of network conditions is supported.
Public Function canEmulateNetworkConditions() AS boolean
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.canEmulateNetworkConditions", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then canEmulateNetworkConditions = results("result")
    End If
End Function

' Clears browser cache.
Public Sub clearBrowserCache()
    Dim params As New Dictionary
    cdp.InvokeMethod "Network.clearBrowserCache", params
End Sub

' Clears browser cookies.
Public Sub clearBrowserCookies()
    Dim params As New Dictionary
    cdp.InvokeMethod "Network.clearBrowserCookies", params
End Sub

' Response to Network.requestIntercepted which either modifies the request to continue with any
' modifications, or blocks it, or completes it with the provided response bytes. If a network
' fetch occurs as a result which encounters a redirect an additional Network.requestIntercepted
' event will be sent with the same InterceptionId.
' Deprecated, use Fetch.continueRequest, Fetch.fulfillRequest and Fetch.failRequest instead.
' Experimental
Public Sub continueInterceptedRequest( _
    ByVal interceptionId AS string, _
    Optional ByVal errorReason AS Variant, _
    Optional ByVal rawResponse AS Variant, _
    Optional ByVal url AS Variant, _
    Optional ByVal method AS Variant, _
    Optional ByVal postData AS Variant, _
    Optional ByVal headers AS Variant, _
    Optional ByVal authChallengeResponse AS Variant _
)
    ' interceptionId: InterceptionId 
    ' errorReason: ErrorReason(optional) If set this causes the request to fail with the given reason. Passing `Aborted` for requests
    '   marked with `isNavigationRequest` also cancels the navigation. Must not be set in response
    '   to an authChallenge.
    ' rawResponse: binary(optional) If set the requests completes using with the provided base64 encoded raw response, including
    '   HTTP status line and headers etc... Must not be set in response to an authChallenge.
    ' url: string(optional) If set the request url will be modified in a way that's not observable by page. Must not be
    '   set in response to an authChallenge.
    ' method: string(optional) If set this allows the request method to be overridden. Must not be set in response to an
    '   authChallenge.
    ' postData: string(optional) If set this allows postData to be set. Must not be set in response to an authChallenge.
    ' headers: Headers(optional) If set this allows the request headers to be changed. Must not be set in response to an
    '   authChallenge.
    ' authChallengeResponse: AuthChallengeResponse(optional) Response to a requestIntercepted with an authChallenge. Must not be set otherwise.

    Dim params As New Dictionary
    params("interceptionId") = CStr(interceptionId)
    If Not IsMissing(errorReason) Then params("errorReason") = CStr(errorReason)
    If Not IsMissing(rawResponse) Then params("rawResponse") = CStr(rawResponse)
    If Not IsMissing(url) Then params("url") = CStr(url)
    If Not IsMissing(method) Then params("method") = CStr(method)
    If Not IsMissing(postData) Then params("postData") = CStr(postData)
    If Not IsMissing(headers) Then Set params("headers") = headers
    If Not IsMissing(authChallengeResponse) Then Set params("authChallengeResponse") = authChallengeResponse

    cdp.InvokeMethod "Network.continueInterceptedRequest", params
End Sub

' Deletes browser cookies with matching name and url or domain/path pair.
Public Sub deleteCookies( _
    ByVal name AS string, _
    Optional ByVal url AS Variant, _
    Optional ByVal domain AS Variant, _
    Optional ByVal path AS Variant _
)
    ' name: string Name of the cookies to remove.
    ' url: string(optional) If specified, deletes all the cookies with the given name where domain and path match
    '   provided URL.
    ' domain: string(optional) If specified, deletes only cookies with the exact domain.
    ' path: string(optional) If specified, deletes only cookies with the exact path.

    Dim params As New Dictionary
    params("name") = CStr(name)
    If Not IsMissing(url) Then params("url") = CStr(url)
    If Not IsMissing(domain) Then params("domain") = CStr(domain)
    If Not IsMissing(path) Then params("path") = CStr(path)

    cdp.InvokeMethod "Network.deleteCookies", params
End Sub

' Disables network tracking, prevents network events from being sent to the client.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Network.disable", params
End Sub

' Activates emulation of network conditions.
Public Sub emulateNetworkConditions( _
    ByVal offline AS boolean, _
    ByVal latency AS Double, _
    ByVal downloadThroughput AS Double, _
    ByVal uploadThroughput AS Double, _
    Optional ByVal connectionType AS Variant _
)
    ' offline: boolean True to emulate internet disconnection.
    ' latency: number Minimum latency from request sent to response headers received (ms).
    ' downloadThroughput: number Maximal aggregated download throughput (bytes/sec). -1 disables download throttling.
    ' uploadThroughput: number Maximal aggregated upload throughput (bytes/sec).  -1 disables upload throttling.
    ' connectionType: ConnectionType(optional) Connection type if known.

    Dim params As New Dictionary
    params("offline") = CBool(offline)
    params("latency") = CDbl(latency)
    params("downloadThroughput") = CDbl(downloadThroughput)
    params("uploadThroughput") = CDbl(uploadThroughput)
    If Not IsMissing(connectionType) Then params("connectionType") = CStr(connectionType)

    cdp.InvokeMethod "Network.emulateNetworkConditions", params
End Sub

' Enables network tracking, network events will now be delivered to the client.
Public Sub enable( _
    Optional ByVal maxTotalBufferSize AS Variant, _
    Optional ByVal maxResourceBufferSize AS Variant, _
    Optional ByVal maxPostDataSize AS Variant _
)
    ' maxTotalBufferSize: integer(optional) Buffer size in bytes to use when preserving network payloads (XHRs, etc).
    ' maxResourceBufferSize: integer(optional) Per-resource buffer size in bytes to use when preserving network payloads (XHRs, etc).
    ' maxPostDataSize: integer(optional) Longest post body size (in bytes) that would be included in requestWillBeSent notification

    Dim params As New Dictionary
    If Not IsMissing(maxTotalBufferSize) Then params("maxTotalBufferSize") = CLng(maxTotalBufferSize)
    If Not IsMissing(maxResourceBufferSize) Then params("maxResourceBufferSize") = CLng(maxResourceBufferSize)
    If Not IsMissing(maxPostDataSize) Then params("maxPostDataSize") = CLng(maxPostDataSize)

    cdp.InvokeMethod "Network.enable", params
End Sub

' Returns all browser cookies. Depending on the backend support, will return detailed cookie
' information in the `cookies` field.
' Returns: cookies - Array of cookie objects.
Public Function getAllCookies() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.getAllCookies", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("cookies") Then Set getAllCookies = results("cookies")
    End If
End Function

' Returns the DER-encoded certificate.
' Returns: tableNames - 
' Experimental
Public Function getCertificate( _
    ByVal origin AS string _
) AS Collection
    ' origin: string Origin to get certificate for.

    Dim params As New Dictionary
    params("origin") = CStr(origin)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.getCertificate", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("tableNames") Then Set getCertificate = results("tableNames")
    End If
End Function

' Returns all browser cookies for the current URL. Depending on the backend support, will return
' detailed cookie information in the `cookies` field.
' Returns: cookies - Array of cookie objects.
Public Function getCookies( _
    Optional ByVal urls AS Variant _
) AS Collection
    ' urls: array(optional) The list of URLs for which applicable cookies will be fetched.
    '   If not specified, it's assumed to be set to the list containing
    '   the URLs of the page and all of its subframes.

    Dim params As New Dictionary
    If Not IsMissing(urls) Then Set params("urls") = urls

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.getCookies", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("cookies") Then Set getCookies = results("cookies")
    End If
End Function

' Returns content served for the given request.
' Returns:  - 
Public Function getResponseBody( _
    ByVal requestId AS string _
) AS Dictionary
    ' requestId: RequestId Identifier of the network request to get content for.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.getResponseBody", params)
    If cdp.ErrorCode = 0 Then
        Set getResponseBody = results
    End If
End Function

' Returns post data sent with the request. Returns an error when no data was sent with the request.
' Returns: postData - Request body string, omitting files from multipart requests
Public Function getRequestPostData( _
    ByVal requestId AS string _
) AS string
    ' requestId: RequestId Identifier of the network request to get content for.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.getRequestPostData", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("postData") Then getRequestPostData = results("postData")
    End If
End Function

' Returns content served for the given currently intercepted request.
' Returns:  - 
' Experimental
Public Function getResponseBodyForInterception( _
    ByVal interceptionId AS string _
) AS Dictionary
    ' interceptionId: InterceptionId Identifier for the intercepted request to get body for.

    Dim params As New Dictionary
    params("interceptionId") = CStr(interceptionId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.getResponseBodyForInterception", params)
    If cdp.ErrorCode = 0 Then
        Set getResponseBodyForInterception = results
    End If
End Function

' Returns a handle to the stream representing the response body. Note that after this command,
' the intercepted request can't be continued as is -- you either need to cancel it or to provide
' the response body. The stream only supports sequential read, IO.read will fail if the position
' is specified.
' Returns: stream - 
' Experimental
Public Function takeResponseBodyForInterceptionAsStream( _
    ByVal interceptionId AS string _
) AS string
    ' interceptionId: InterceptionId 

    Dim params As New Dictionary
    params("interceptionId") = CStr(interceptionId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.takeResponseBodyForInterceptionAsStream", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("stream") Then takeResponseBodyForInterceptionAsStream = results("stream")
    End If
End Function

' This method sends a new XMLHttpRequest which is identical to the original one. The following
' parameters should be identical: method, url, async, request body, extra headers, withCredentials
' attribute, user, password.
' Experimental
Public Sub replayXHR( _
    ByVal requestId AS string _
)
    ' requestId: RequestId Identifier of XHR to replay.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)

    cdp.InvokeMethod "Network.replayXHR", params
End Sub

' Searches for given string in response content.
' Returns: result - List of search matches.
' Experimental
Public Function searchInResponseBody( _
    ByVal requestId AS string, _
    ByVal query AS string, _
    Optional ByVal caseSensitive AS Variant, _
    Optional ByVal isRegex AS Variant _
) AS Collection
    ' requestId: RequestId Identifier of the network response to search.
    ' query: string String to search for.
    ' caseSensitive: boolean(optional) If true, search is case sensitive.
    ' isRegex: boolean(optional) If true, treats string parameter as regex.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)
    params("query") = CStr(query)
    If Not IsMissing(caseSensitive) Then params("caseSensitive") = CBool(caseSensitive)
    If Not IsMissing(isRegex) Then params("isRegex") = CBool(isRegex)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.searchInResponseBody", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then Set searchInResponseBody = results("result")
    End If
End Function

' Blocks URLs from loading.
' Experimental
Public Sub setBlockedURLs( _
    ByVal urls AS Collection _
)
    ' urls: array URL patterns to block. Wildcards ('*') are allowed.

    Dim params As New Dictionary
    Set params("urls") = urls

    cdp.InvokeMethod "Network.setBlockedURLs", params
End Sub

' Toggles ignoring of service worker for each request.
' Experimental
Public Sub setBypassServiceWorker( _
    ByVal bypass AS boolean _
)
    ' bypass: boolean Bypass service worker and load from network.

    Dim params As New Dictionary
    params("bypass") = CBool(bypass)

    cdp.InvokeMethod "Network.setBypassServiceWorker", params
End Sub

' Toggles ignoring cache for each request. If `true`, cache will not be used.
Public Sub setCacheDisabled( _
    ByVal cacheDisabled AS boolean _
)
    ' cacheDisabled: boolean Cache disabled state.

    Dim params As New Dictionary
    params("cacheDisabled") = CBool(cacheDisabled)

    cdp.InvokeMethod "Network.setCacheDisabled", params
End Sub

' Sets a cookie with the given cookie data; may overwrite equivalent cookies if they exist.
' Returns: success - Always set to true. If an error occurs, the response indicates protocol error.
Public Function setCookie( _
    ByVal name AS string, _
    ByVal value AS string, _
    Optional ByVal url AS Variant, _
    Optional ByVal domain AS Variant, _
    Optional ByVal path AS Variant, _
    Optional ByVal secure AS Variant, _
    Optional ByVal httpOnly AS Variant, _
    Optional ByVal sameSite AS Variant, _
    Optional ByVal expires AS Variant, _
    Optional ByVal priority AS Variant, _
    Optional ByVal sameParty AS Variant, _
    Optional ByVal sourceScheme AS Variant, _
    Optional ByVal sourcePort AS Variant _
) AS boolean
    ' name: string Cookie name.
    ' value: string Cookie value.
    ' url: string(optional) The request-URI to associate with the setting of the cookie. This value can affect the
    '   default domain, path, source port, and source scheme values of the created cookie.
    ' domain: string(optional) Cookie domain.
    ' path: string(optional) Cookie path.
    ' secure: boolean(optional) True if cookie is secure.
    ' httpOnly: boolean(optional) True if cookie is http-only.
    ' sameSite: CookieSameSite(optional) Cookie SameSite type.
    ' expires: TimeSinceEpoch(optional) Cookie expiration date, session cookie if not set
    ' priority: CookiePriority(optional) Cookie Priority type.
    ' sameParty: boolean(optional) True if cookie is SameParty.
    ' sourceScheme: CookieSourceScheme(optional) Cookie source scheme type.
    ' sourcePort: integer(optional) Cookie source port. Valid values are {-1, [1, 65535]}, -1 indicates an unspecified port.
    '   An unspecified port value allows protocol clients to emulate legacy cookie scope for the port.
    '   This is a temporary ability and it will be removed in the future.

    Dim params As New Dictionary
    params("name") = CStr(name)
    params("value") = CStr(value)
    If Not IsMissing(url) Then params("url") = CStr(url)
    If Not IsMissing(domain) Then params("domain") = CStr(domain)
    If Not IsMissing(path) Then params("path") = CStr(path)
    If Not IsMissing(secure) Then params("secure") = CBool(secure)
    If Not IsMissing(httpOnly) Then params("httpOnly") = CBool(httpOnly)
    If Not IsMissing(sameSite) Then params("sameSite") = CStr(sameSite)
    If Not IsMissing(expires) Then params("expires") = CDbl(expires)
    If Not IsMissing(priority) Then params("priority") = CStr(priority)
    If Not IsMissing(sameParty) Then params("sameParty") = CBool(sameParty)
    If Not IsMissing(sourceScheme) Then params("sourceScheme") = CStr(sourceScheme)
    If Not IsMissing(sourcePort) Then params("sourcePort") = CLng(sourcePort)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.setCookie", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("success") Then setCookie = results("success")
    End If
End Function

' Sets given cookies.
Public Sub setCookies( _
    ByVal cookies AS Collection _
)
    ' cookies: array Cookies to be set.

    Dim params As New Dictionary
    Set params("cookies") = cookies

    cdp.InvokeMethod "Network.setCookies", params
End Sub

' Specifies whether to always send extra HTTP headers with the requests from this page.
Public Sub setExtraHTTPHeaders( _
    ByVal headers AS Dictionary _
)
    ' headers: Headers Map with extra HTTP headers.

    Dim params As New Dictionary
    Set params("headers") = headers

    cdp.InvokeMethod "Network.setExtraHTTPHeaders", params
End Sub

' Specifies whether to attach a page script stack id in requests
' Experimental
Public Sub setAttachDebugStack( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean Whether to attach a page script stack for debugging purpose.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Network.setAttachDebugStack", params
End Sub

' Sets the requests to intercept that match the provided patterns and optionally resource types.
' Deprecated, please use Fetch.enable instead.
' Experimental
Public Sub setRequestInterception( _
    ByVal patterns AS Collection _
)
    ' patterns: array Requests matching any of these patterns will be forwarded and wait for the corresponding
    '   continueInterceptedRequest call.

    Dim params As New Dictionary
    Set params("patterns") = patterns

    cdp.InvokeMethod "Network.setRequestInterception", params
End Sub

' Allows overriding user agent with the given string.
Public Sub setUserAgentOverride( _
    ByVal userAgent AS string, _
    Optional ByVal acceptLanguage AS Variant, _
    Optional ByVal platform AS Variant, _
    Optional ByVal userAgentMetadata AS Variant _
)
    ' userAgent: string User agent to use.
    ' acceptLanguage: string(optional) Browser langugage to emulate.
    ' platform: string(optional) The platform navigator.platform should return.
    ' userAgentMetadata: Emulation.UserAgentMetadata(optional) To be sent in Sec-CH-UA-* headers and returned in navigator.userAgentData

    Dim params As New Dictionary
    params("userAgent") = CStr(userAgent)
    If Not IsMissing(acceptLanguage) Then params("acceptLanguage") = CStr(acceptLanguage)
    If Not IsMissing(platform) Then params("platform") = CStr(platform)
    If Not IsMissing(userAgentMetadata) Then Set params("userAgentMetadata") = userAgentMetadata

    cdp.InvokeMethod "Network.setUserAgentOverride", params
End Sub

' Returns information about the COEP/COOP isolation status.
' Returns: status - 
' Experimental
Public Function getSecurityIsolationStatus( _
    Optional ByVal frameId AS Variant _
) AS Dictionary
    ' frameId: Page.FrameId(optional) If no frameId is provided, the status of the target is provided.

    Dim params As New Dictionary
    If Not IsMissing(frameId) Then params("frameId") = CStr(frameId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.getSecurityIsolationStatus", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("status") Then Set getSecurityIsolationStatus = results("status")
    End If
End Function

' Enables tracking for the Reporting API, events generated by the Reporting API will now be delivered to the client.
' Enabling triggers 'reportingApiReportAdded' for all existing reports.
' Experimental
Public Sub enableReportingApi( _
    ByVal enable AS boolean _
)
    ' enable: boolean Whether to enable or disable events for the Reporting API

    Dim params As New Dictionary
    params("enable") = CBool(enable)

    cdp.InvokeMethod "Network.enableReportingApi", params
End Sub

' Fetches the resource and returns the content.
' Returns: resource - 
' Experimental
Public Function loadNetworkResource( _
    ByVal url AS string, _
    ByVal options AS Dictionary, _
    Optional ByVal frameId AS Variant _
) AS Dictionary
    ' frameId: Page.FrameId(optional) Frame id to get the resource for. Mandatory for frame targets, and
    '   should be omitted for worker targets.
    ' url: string URL of the resource to get content for.
    ' options: LoadNetworkResourceOptions Options for the request.

    Dim params As New Dictionary
    If Not IsMissing(frameId) Then params("frameId") = CStr(frameId)
    params("url") = CStr(url)
    Set params("options") = options

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Network.loadNetworkResource", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("resource") Then Set loadNetworkResource = results("resource")
    End If
End Function

