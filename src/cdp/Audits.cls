VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAudits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpAudits"
' Audits [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Information about a cookie that is affected by an inspector issue.
' AffectedCookie AS object

' Information about a request that is affected by an inspector issue.
' AffectedRequest AS object

' Information about the frame affected by an inspector issue.
' AffectedFrame AS object

' 
' SameSiteCookieExclusionReason AS string
Public Enum SameSiteCookieExclusionReason
    SSCER_ExcludeSameSiteUnspecifiedTreatedAsLax
    SSCER_ExcludeSameSiteNoneInsecure
    SSCER_ExcludeSameSiteLax
    SSCER_ExcludeSameSiteStrict
    SSCER_ExcludeInvalidSameParty
    SSCER_ExcludeSamePartyCrossPartyContext
End Enum

' 
' SameSiteCookieWarningReason AS string
Public Enum SameSiteCookieWarningReason
    SSCWR_WarnSameSiteUnspecifiedCrossSiteContext
    SSCWR_WarnSameSiteNoneInsecure
    SSCWR_WarnSameSiteUnspecifiedLaxAllowUnsafe
    SSCWR_WarnSameSiteStrictLaxDowngradeStrict
    SSCWR_WarnSameSiteStrictCrossDowngradeStrict
    SSCWR_WarnSameSiteStrictCrossDowngradeLax
    SSCWR_WarnSameSiteLaxCrossDowngradeStrict
    SSCWR_WarnSameSiteLaxCrossDowngradeLax
End Enum

' 
' SameSiteCookieOperation AS string
Public Enum SameSiteCookieOperation
    SSCO_SetCookie
    SSCO_ReadCookie
End Enum

' This information is currently necessary, as the front-end has a difficult
'   time finding a specific cookie. With this, we can convey specific error
'   information without the cookie.
' SameSiteCookieIssueDetails AS object

' 
' MixedContentResolutionStatus AS string
Public Enum MixedContentResolutionStatus
    MCRS_MixedContentBlocked
    MCRS_MixedContentAutomaticallyUpgraded
    MCRS_MixedContentWarning
End Enum

' 
' MixedContentResourceType AS string
Public Enum MixedContentResourceType
    MCRT_Audio
    MCRT_Beacon
    MCRT_CSPReport
    MCRT_Download
    MCRT_EventSource
    MCRT_Favicon
    MCRT_Font
    MCRT_Form
    MCRT_Frame
    MCRT_Image
    MCRT_Import
    MCRT_Manifest
    MCRT_Ping
    MCRT_PluginData
    MCRT_PluginResource
    MCRT_Prefetch
    MCRT_Resource
    MCRT_Script
    MCRT_ServiceWorker
    MCRT_SharedWorker
    MCRT_Stylesheet
    MCRT_Track
    MCRT_Video
    MCRT_Worker
    MCRT_XMLHttpRequest
    MCRT_XSLT
End Enum

' 
' MixedContentIssueDetails AS object

' Enum indicating the reason a response has been blocked. These reasons are
'   refinements of the net error BLOCKED_BY_RESPONSE.
' BlockedByResponseReason AS string
Public Enum BlockedByResponseReason
    BBRR_CoepFrameResourceNeedsCoepHeader
    BBRR_CoopSandboxedIFrameCannotNavigateToCoopPage
    BBRR_CorpNotSameOrigin
    BBRR_CorpNotSameOriginAfterDefaultedToSameOriginByCoep
    BBRR_CorpNotSameSite
End Enum

' Details for a request that has been blocked with the BLOCKED_BY_RESPONSE
'   code. Currently only used for COEP/COOP, but may be extended to include
'   some CSP errors in the future.
' BlockedByResponseIssueDetails AS object

' 
' HeavyAdResolutionStatus AS string
Public Enum HeavyAdResolutionStatus
    HARS_HeavyAdBlocked
    HARS_HeavyAdWarning
End Enum

' 
' HeavyAdReason AS string
Public Enum HeavyAdReason
    HAR_NetworkTotalLimit
    HAR_CpuTotalLimit
    HAR_CpuPeakLimit
End Enum

' 
' HeavyAdIssueDetails AS object

' 
' ContentSecurityPolicyViolationType AS string
Public Enum ContentSecurityPolicyViolationType
    CSPVT_kInlineViolation
    CSPVT_kEvalViolation
    CSPVT_kURLViolation
    CSPVT_kTrustedTypesSinkViolation
    CSPVT_kTrustedTypesPolicyViolation
End Enum

' 
' SourceCodeLocation AS object

' 
' ContentSecurityPolicyIssueDetails AS object

' 
' SharedArrayBufferIssueType AS string
Public Enum SharedArrayBufferIssueType
    SABIT_TransferIssue
    SABIT_CreationIssue
End Enum

' Details for a issue arising from an SAB being instantiated in, or
'   transferred to a context that is not cross-origin isolated.
' SharedArrayBufferIssueDetails AS object

' 
' TwaQualityEnforcementViolationType AS string
Public Enum TwaQualityEnforcementViolationType
    TQEVT_kHttpError
    TQEVT_kUnavailableOffline
    TQEVT_kDigitalAssetLinks
End Enum

' 
' TrustedWebActivityIssueDetails AS object

' 
' LowTextContrastIssueDetails AS object

' Details for a CORS related issue, e.g. a warning or error related to
'   CORS RFC1918 enforcement.
' CorsIssueDetails AS object

' 
' AttributionReportingIssueType AS string
Public Enum AttributionReportingIssueType
    ARIT_PermissionPolicyDisabled
    ARIT_InvalidAttributionSourceEventId
    ARIT_InvalidAttributionData
    ARIT_AttributionSourceUntrustworthyOrigin
    ARIT_AttributionUntrustworthyOrigin
    ARIT_AttributionTriggerDataTooLarge
    ARIT_AttributionEventSourceTriggerDataTooLarge
End Enum

' Details for issues around "Attribution Reporting API" usage.
'   Explainer: https://github.com/WICG/conversion-measurement-api
' AttributionReportingIssueDetails AS object

' Details for issues about documents in Quirks Mode
'   or Limited Quirks Mode that affects page layouting.
' QuirksModeIssueDetails AS object

' 
' NavigatorUserAgentIssueDetails AS object

' 
' WasmCrossOriginModuleSharingIssueDetails AS object

' 
' GenericIssueErrorType AS string
Public Enum GenericIssueErrorType
    GIET_CrossOriginPortalPostMessageError
End Enum

' Depending on the concrete errorType, different properties are set.
' GenericIssueDetails AS object

' A unique identifier for the type of issue. Each type may use one of the
'   optional fields in InspectorIssueDetails to convey more specific
'   information about the kind of issue.
' InspectorIssueCode AS string
Public Enum InspectorIssueCode
    IIC_SameSiteCookieIssue
    IIC_MixedContentIssue
    IIC_BlockedByResponseIssue
    IIC_HeavyAdIssue
    IIC_ContentSecurityPolicyIssue
    IIC_SharedArrayBufferIssue
    IIC_TrustedWebActivityIssue
    IIC_LowTextContrastIssue
    IIC_CorsIssue
    IIC_AttributionReportingIssue
    IIC_QuirksModeIssue
    IIC_NavigatorUserAgentIssue
    IIC_WasmCrossOriginModuleSharingIssue
    IIC_GenericIssue
End Enum

' This struct holds a list of optional fields with additional information
'   specific to the kind of issue. When adding a new issue code, please also
'   add a new optional field to this type.
' InspectorIssueDetails AS object

' A unique id for a DevTools inspector issue. Allows other entities (e.g.
'   exceptions, CDP message, console messages, etc.) to reference an issue.
' IssueId AS string

' An inspector issue reported from the back-end.
' InspectorIssue AS object


Public Function SameSiteCookieExclusionReasonToString(ByVal sscer As SameSiteCookieExclusionReason) As String
    Dim retVal As String
    Select Case sscer
        Case SSCER_ExcludeSameSiteUnspecifiedTreatedAsLax
            retVal = "ExcludeSameSiteUnspecifiedTreatedAsLax"
        Case SSCER_ExcludeSameSiteNoneInsecure
            retVal = "ExcludeSameSiteNoneInsecure"
        Case SSCER_ExcludeSameSiteLax
            retVal = "ExcludeSameSiteLax"
        Case SSCER_ExcludeSameSiteStrict
            retVal = "ExcludeSameSiteStrict"
        Case SSCER_ExcludeInvalidSameParty
            retVal = "ExcludeInvalidSameParty"
        Case SSCER_ExcludeSamePartyCrossPartyContext
            retVal = "ExcludeSamePartyCrossPartyContext"
        Case Else
            Debug.Print "Warning, unknown value " & sscer
    End Select
    SameSiteCookieExclusionReasonToString = retVal
End Function

Public Function StringToSameSiteCookieExclusionReason(ByVal s As String) As SameSiteCookieExclusionReason
    Dim retVal As SameSiteCookieExclusionReason
    Select Case s
        Case "ExcludeSameSiteUnspecifiedTreatedAsLax"
            retVal = SSCER_ExcludeSameSiteUnspecifiedTreatedAsLax
        Case "ExcludeSameSiteNoneInsecure"
            retVal = SSCER_ExcludeSameSiteNoneInsecure
        Case "ExcludeSameSiteLax"
            retVal = SSCER_ExcludeSameSiteLax
        Case "ExcludeSameSiteStrict"
            retVal = SSCER_ExcludeSameSiteStrict
        Case "ExcludeInvalidSameParty"
            retVal = SSCER_ExcludeInvalidSameParty
        Case "ExcludeSamePartyCrossPartyContext"
            retVal = SSCER_ExcludeSamePartyCrossPartyContext
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSameSiteCookieExclusionReason = retVal
End Function


Public Function SameSiteCookieWarningReasonToString(ByVal sscwr As SameSiteCookieWarningReason) As String
    Dim retVal As String
    Select Case sscwr
        Case SSCWR_WarnSameSiteUnspecifiedCrossSiteContext
            retVal = "WarnSameSiteUnspecifiedCrossSiteContext"
        Case SSCWR_WarnSameSiteNoneInsecure
            retVal = "WarnSameSiteNoneInsecure"
        Case SSCWR_WarnSameSiteUnspecifiedLaxAllowUnsafe
            retVal = "WarnSameSiteUnspecifiedLaxAllowUnsafe"
        Case SSCWR_WarnSameSiteStrictLaxDowngradeStrict
            retVal = "WarnSameSiteStrictLaxDowngradeStrict"
        Case SSCWR_WarnSameSiteStrictCrossDowngradeStrict
            retVal = "WarnSameSiteStrictCrossDowngradeStrict"
        Case SSCWR_WarnSameSiteStrictCrossDowngradeLax
            retVal = "WarnSameSiteStrictCrossDowngradeLax"
        Case SSCWR_WarnSameSiteLaxCrossDowngradeStrict
            retVal = "WarnSameSiteLaxCrossDowngradeStrict"
        Case SSCWR_WarnSameSiteLaxCrossDowngradeLax
            retVal = "WarnSameSiteLaxCrossDowngradeLax"
        Case Else
            Debug.Print "Warning, unknown value " & sscwr
    End Select
    SameSiteCookieWarningReasonToString = retVal
End Function

Public Function StringToSameSiteCookieWarningReason(ByVal s As String) As SameSiteCookieWarningReason
    Dim retVal As SameSiteCookieWarningReason
    Select Case s
        Case "WarnSameSiteUnspecifiedCrossSiteContext"
            retVal = SSCWR_WarnSameSiteUnspecifiedCrossSiteContext
        Case "WarnSameSiteNoneInsecure"
            retVal = SSCWR_WarnSameSiteNoneInsecure
        Case "WarnSameSiteUnspecifiedLaxAllowUnsafe"
            retVal = SSCWR_WarnSameSiteUnspecifiedLaxAllowUnsafe
        Case "WarnSameSiteStrictLaxDowngradeStrict"
            retVal = SSCWR_WarnSameSiteStrictLaxDowngradeStrict
        Case "WarnSameSiteStrictCrossDowngradeStrict"
            retVal = SSCWR_WarnSameSiteStrictCrossDowngradeStrict
        Case "WarnSameSiteStrictCrossDowngradeLax"
            retVal = SSCWR_WarnSameSiteStrictCrossDowngradeLax
        Case "WarnSameSiteLaxCrossDowngradeStrict"
            retVal = SSCWR_WarnSameSiteLaxCrossDowngradeStrict
        Case "WarnSameSiteLaxCrossDowngradeLax"
            retVal = SSCWR_WarnSameSiteLaxCrossDowngradeLax
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSameSiteCookieWarningReason = retVal
End Function


Public Function SameSiteCookieOperationToString(ByVal ssco As SameSiteCookieOperation) As String
    Dim retVal As String
    Select Case ssco
        Case SSCO_SetCookie
            retVal = "SetCookie"
        Case SSCO_ReadCookie
            retVal = "ReadCookie"
        Case Else
            Debug.Print "Warning, unknown value " & ssco
    End Select
    SameSiteCookieOperationToString = retVal
End Function

Public Function StringToSameSiteCookieOperation(ByVal s As String) As SameSiteCookieOperation
    Dim retVal As SameSiteCookieOperation
    Select Case s
        Case "SetCookie"
            retVal = SSCO_SetCookie
        Case "ReadCookie"
            retVal = SSCO_ReadCookie
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSameSiteCookieOperation = retVal
End Function


Public Function MixedContentResolutionStatusToString(ByVal mcrs As MixedContentResolutionStatus) As String
    Dim retVal As String
    Select Case mcrs
        Case MCRS_MixedContentBlocked
            retVal = "MixedContentBlocked"
        Case MCRS_MixedContentAutomaticallyUpgraded
            retVal = "MixedContentAutomaticallyUpgraded"
        Case MCRS_MixedContentWarning
            retVal = "MixedContentWarning"
        Case Else
            Debug.Print "Warning, unknown value " & mcrs
    End Select
    MixedContentResolutionStatusToString = retVal
End Function

Public Function StringToMixedContentResolutionStatus(ByVal s As String) As MixedContentResolutionStatus
    Dim retVal As MixedContentResolutionStatus
    Select Case s
        Case "MixedContentBlocked"
            retVal = MCRS_MixedContentBlocked
        Case "MixedContentAutomaticallyUpgraded"
            retVal = MCRS_MixedContentAutomaticallyUpgraded
        Case "MixedContentWarning"
            retVal = MCRS_MixedContentWarning
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToMixedContentResolutionStatus = retVal
End Function


Public Function MixedContentResourceTypeToString(ByVal mcrt As MixedContentResourceType) As String
    Dim retVal As String
    Select Case mcrt
        Case MCRT_Audio
            retVal = "Audio"
        Case MCRT_Beacon
            retVal = "Beacon"
        Case MCRT_CSPReport
            retVal = "CSPReport"
        Case MCRT_Download
            retVal = "Download"
        Case MCRT_EventSource
            retVal = "EventSource"
        Case MCRT_Favicon
            retVal = "Favicon"
        Case MCRT_Font
            retVal = "Font"
        Case MCRT_Form
            retVal = "Form"
        Case MCRT_Frame
            retVal = "Frame"
        Case MCRT_Image
            retVal = "Image"
        Case MCRT_Import
            retVal = "Import"
        Case MCRT_Manifest
            retVal = "Manifest"
        Case MCRT_Ping
            retVal = "Ping"
        Case MCRT_PluginData
            retVal = "PluginData"
        Case MCRT_PluginResource
            retVal = "PluginResource"
        Case MCRT_Prefetch
            retVal = "Prefetch"
        Case MCRT_Resource
            retVal = "Resource"
        Case MCRT_Script
            retVal = "Script"
        Case MCRT_ServiceWorker
            retVal = "ServiceWorker"
        Case MCRT_SharedWorker
            retVal = "SharedWorker"
        Case MCRT_Stylesheet
            retVal = "Stylesheet"
        Case MCRT_Track
            retVal = "Track"
        Case MCRT_Video
            retVal = "Video"
        Case MCRT_Worker
            retVal = "Worker"
        Case MCRT_XMLHttpRequest
            retVal = "XMLHttpRequest"
        Case MCRT_XSLT
            retVal = "XSLT"
        Case Else
            Debug.Print "Warning, unknown value " & mcrt
    End Select
    MixedContentResourceTypeToString = retVal
End Function

Public Function StringToMixedContentResourceType(ByVal s As String) As MixedContentResourceType
    Dim retVal As MixedContentResourceType
    Select Case s
        Case "Audio"
            retVal = MCRT_Audio
        Case "Beacon"
            retVal = MCRT_Beacon
        Case "CSPReport"
            retVal = MCRT_CSPReport
        Case "Download"
            retVal = MCRT_Download
        Case "EventSource"
            retVal = MCRT_EventSource
        Case "Favicon"
            retVal = MCRT_Favicon
        Case "Font"
            retVal = MCRT_Font
        Case "Form"
            retVal = MCRT_Form
        Case "Frame"
            retVal = MCRT_Frame
        Case "Image"
            retVal = MCRT_Image
        Case "Import"
            retVal = MCRT_Import
        Case "Manifest"
            retVal = MCRT_Manifest
        Case "Ping"
            retVal = MCRT_Ping
        Case "PluginData"
            retVal = MCRT_PluginData
        Case "PluginResource"
            retVal = MCRT_PluginResource
        Case "Prefetch"
            retVal = MCRT_Prefetch
        Case "Resource"
            retVal = MCRT_Resource
        Case "Script"
            retVal = MCRT_Script
        Case "ServiceWorker"
            retVal = MCRT_ServiceWorker
        Case "SharedWorker"
            retVal = MCRT_SharedWorker
        Case "Stylesheet"
            retVal = MCRT_Stylesheet
        Case "Track"
            retVal = MCRT_Track
        Case "Video"
            retVal = MCRT_Video
        Case "Worker"
            retVal = MCRT_Worker
        Case "XMLHttpRequest"
            retVal = MCRT_XMLHttpRequest
        Case "XSLT"
            retVal = MCRT_XSLT
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToMixedContentResourceType = retVal
End Function


Public Function BlockedByResponseReasonToString(ByVal bbrr As BlockedByResponseReason) As String
    Dim retVal As String
    Select Case bbrr
        Case BBRR_CoepFrameResourceNeedsCoepHeader
            retVal = "CoepFrameResourceNeedsCoepHeader"
        Case BBRR_CoopSandboxedIFrameCannotNavigateToCoopPage
            retVal = "CoopSandboxedIFrameCannotNavigateToCoopPage"
        Case BBRR_CorpNotSameOrigin
            retVal = "CorpNotSameOrigin"
        Case BBRR_CorpNotSameOriginAfterDefaultedToSameOriginByCoep
            retVal = "CorpNotSameOriginAfterDefaultedToSameOriginByCoep"
        Case BBRR_CorpNotSameSite
            retVal = "CorpNotSameSite"
        Case Else
            Debug.Print "Warning, unknown value " & bbrr
    End Select
    BlockedByResponseReasonToString = retVal
End Function

Public Function StringToBlockedByResponseReason(ByVal s As String) As BlockedByResponseReason
    Dim retVal As BlockedByResponseReason
    Select Case s
        Case "CoepFrameResourceNeedsCoepHeader"
            retVal = BBRR_CoepFrameResourceNeedsCoepHeader
        Case "CoopSandboxedIFrameCannotNavigateToCoopPage"
            retVal = BBRR_CoopSandboxedIFrameCannotNavigateToCoopPage
        Case "CorpNotSameOrigin"
            retVal = BBRR_CorpNotSameOrigin
        Case "CorpNotSameOriginAfterDefaultedToSameOriginByCoep"
            retVal = BBRR_CorpNotSameOriginAfterDefaultedToSameOriginByCoep
        Case "CorpNotSameSite"
            retVal = BBRR_CorpNotSameSite
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToBlockedByResponseReason = retVal
End Function


Public Function HeavyAdResolutionStatusToString(ByVal hars As HeavyAdResolutionStatus) As String
    Dim retVal As String
    Select Case hars
        Case HARS_HeavyAdBlocked
            retVal = "HeavyAdBlocked"
        Case HARS_HeavyAdWarning
            retVal = "HeavyAdWarning"
        Case Else
            Debug.Print "Warning, unknown value " & hars
    End Select
    HeavyAdResolutionStatusToString = retVal
End Function

Public Function StringToHeavyAdResolutionStatus(ByVal s As String) As HeavyAdResolutionStatus
    Dim retVal As HeavyAdResolutionStatus
    Select Case s
        Case "HeavyAdBlocked"
            retVal = HARS_HeavyAdBlocked
        Case "HeavyAdWarning"
            retVal = HARS_HeavyAdWarning
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToHeavyAdResolutionStatus = retVal
End Function


Public Function HeavyAdReasonToString(ByVal har As HeavyAdReason) As String
    Dim retVal As String
    Select Case har
        Case HAR_NetworkTotalLimit
            retVal = "NetworkTotalLimit"
        Case HAR_CpuTotalLimit
            retVal = "CpuTotalLimit"
        Case HAR_CpuPeakLimit
            retVal = "CpuPeakLimit"
        Case Else
            Debug.Print "Warning, unknown value " & har
    End Select
    HeavyAdReasonToString = retVal
End Function

Public Function StringToHeavyAdReason(ByVal s As String) As HeavyAdReason
    Dim retVal As HeavyAdReason
    Select Case s
        Case "NetworkTotalLimit"
            retVal = HAR_NetworkTotalLimit
        Case "CpuTotalLimit"
            retVal = HAR_CpuTotalLimit
        Case "CpuPeakLimit"
            retVal = HAR_CpuPeakLimit
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToHeavyAdReason = retVal
End Function


Public Function ContentSecurityPolicyViolationTypeToString(ByVal cspvt As ContentSecurityPolicyViolationType) As String
    Dim retVal As String
    Select Case cspvt
        Case CSPVT_kInlineViolation
            retVal = "kInlineViolation"
        Case CSPVT_kEvalViolation
            retVal = "kEvalViolation"
        Case CSPVT_kURLViolation
            retVal = "kURLViolation"
        Case CSPVT_kTrustedTypesSinkViolation
            retVal = "kTrustedTypesSinkViolation"
        Case CSPVT_kTrustedTypesPolicyViolation
            retVal = "kTrustedTypesPolicyViolation"
        Case Else
            Debug.Print "Warning, unknown value " & cspvt
    End Select
    ContentSecurityPolicyViolationTypeToString = retVal
End Function

Public Function StringToContentSecurityPolicyViolationType(ByVal s As String) As ContentSecurityPolicyViolationType
    Dim retVal As ContentSecurityPolicyViolationType
    Select Case s
        Case "kInlineViolation"
            retVal = CSPVT_kInlineViolation
        Case "kEvalViolation"
            retVal = CSPVT_kEvalViolation
        Case "kURLViolation"
            retVal = CSPVT_kURLViolation
        Case "kTrustedTypesSinkViolation"
            retVal = CSPVT_kTrustedTypesSinkViolation
        Case "kTrustedTypesPolicyViolation"
            retVal = CSPVT_kTrustedTypesPolicyViolation
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToContentSecurityPolicyViolationType = retVal
End Function


Public Function SharedArrayBufferIssueTypeToString(ByVal sabit As SharedArrayBufferIssueType) As String
    Dim retVal As String
    Select Case sabit
        Case SABIT_TransferIssue
            retVal = "TransferIssue"
        Case SABIT_CreationIssue
            retVal = "CreationIssue"
        Case Else
            Debug.Print "Warning, unknown value " & sabit
    End Select
    SharedArrayBufferIssueTypeToString = retVal
End Function

Public Function StringToSharedArrayBufferIssueType(ByVal s As String) As SharedArrayBufferIssueType
    Dim retVal As SharedArrayBufferIssueType
    Select Case s
        Case "TransferIssue"
            retVal = SABIT_TransferIssue
        Case "CreationIssue"
            retVal = SABIT_CreationIssue
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSharedArrayBufferIssueType = retVal
End Function


Public Function TwaQualityEnforcementViolationTypeToString(ByVal tqevt As TwaQualityEnforcementViolationType) As String
    Dim retVal As String
    Select Case tqevt
        Case TQEVT_kHttpError
            retVal = "kHttpError"
        Case TQEVT_kUnavailableOffline
            retVal = "kUnavailableOffline"
        Case TQEVT_kDigitalAssetLinks
            retVal = "kDigitalAssetLinks"
        Case Else
            Debug.Print "Warning, unknown value " & tqevt
    End Select
    TwaQualityEnforcementViolationTypeToString = retVal
End Function

Public Function StringToTwaQualityEnforcementViolationType(ByVal s As String) As TwaQualityEnforcementViolationType
    Dim retVal As TwaQualityEnforcementViolationType
    Select Case s
        Case "kHttpError"
            retVal = TQEVT_kHttpError
        Case "kUnavailableOffline"
            retVal = TQEVT_kUnavailableOffline
        Case "kDigitalAssetLinks"
            retVal = TQEVT_kDigitalAssetLinks
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToTwaQualityEnforcementViolationType = retVal
End Function


Public Function AttributionReportingIssueTypeToString(ByVal arit As AttributionReportingIssueType) As String
    Dim retVal As String
    Select Case arit
        Case ARIT_PermissionPolicyDisabled
            retVal = "PermissionPolicyDisabled"
        Case ARIT_InvalidAttributionSourceEventId
            retVal = "InvalidAttributionSourceEventId"
        Case ARIT_InvalidAttributionData
            retVal = "InvalidAttributionData"
        Case ARIT_AttributionSourceUntrustworthyOrigin
            retVal = "AttributionSourceUntrustworthyOrigin"
        Case ARIT_AttributionUntrustworthyOrigin
            retVal = "AttributionUntrustworthyOrigin"
        Case ARIT_AttributionTriggerDataTooLarge
            retVal = "AttributionTriggerDataTooLarge"
        Case ARIT_AttributionEventSourceTriggerDataTooLarge
            retVal = "AttributionEventSourceTriggerDataTooLarge"
        Case Else
            Debug.Print "Warning, unknown value " & arit
    End Select
    AttributionReportingIssueTypeToString = retVal
End Function

Public Function StringToAttributionReportingIssueType(ByVal s As String) As AttributionReportingIssueType
    Dim retVal As AttributionReportingIssueType
    Select Case s
        Case "PermissionPolicyDisabled"
            retVal = ARIT_PermissionPolicyDisabled
        Case "InvalidAttributionSourceEventId"
            retVal = ARIT_InvalidAttributionSourceEventId
        Case "InvalidAttributionData"
            retVal = ARIT_InvalidAttributionData
        Case "AttributionSourceUntrustworthyOrigin"
            retVal = ARIT_AttributionSourceUntrustworthyOrigin
        Case "AttributionUntrustworthyOrigin"
            retVal = ARIT_AttributionUntrustworthyOrigin
        Case "AttributionTriggerDataTooLarge"
            retVal = ARIT_AttributionTriggerDataTooLarge
        Case "AttributionEventSourceTriggerDataTooLarge"
            retVal = ARIT_AttributionEventSourceTriggerDataTooLarge
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAttributionReportingIssueType = retVal
End Function


Public Function GenericIssueErrorTypeToString(ByVal giet As GenericIssueErrorType) As String
    Dim retVal As String
    Select Case giet
        Case GIET_CrossOriginPortalPostMessageError
            retVal = "CrossOriginPortalPostMessageError"
        Case Else
            Debug.Print "Warning, unknown value " & giet
    End Select
    GenericIssueErrorTypeToString = retVal
End Function

Public Function StringToGenericIssueErrorType(ByVal s As String) As GenericIssueErrorType
    Dim retVal As GenericIssueErrorType
    Select Case s
        Case "CrossOriginPortalPostMessageError"
            retVal = GIET_CrossOriginPortalPostMessageError
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToGenericIssueErrorType = retVal
End Function


Public Function InspectorIssueCodeToString(ByVal iic As InspectorIssueCode) As String
    Dim retVal As String
    Select Case iic
        Case IIC_SameSiteCookieIssue
            retVal = "SameSiteCookieIssue"
        Case IIC_MixedContentIssue
            retVal = "MixedContentIssue"
        Case IIC_BlockedByResponseIssue
            retVal = "BlockedByResponseIssue"
        Case IIC_HeavyAdIssue
            retVal = "HeavyAdIssue"
        Case IIC_ContentSecurityPolicyIssue
            retVal = "ContentSecurityPolicyIssue"
        Case IIC_SharedArrayBufferIssue
            retVal = "SharedArrayBufferIssue"
        Case IIC_TrustedWebActivityIssue
            retVal = "TrustedWebActivityIssue"
        Case IIC_LowTextContrastIssue
            retVal = "LowTextContrastIssue"
        Case IIC_CorsIssue
            retVal = "CorsIssue"
        Case IIC_AttributionReportingIssue
            retVal = "AttributionReportingIssue"
        Case IIC_QuirksModeIssue
            retVal = "QuirksModeIssue"
        Case IIC_NavigatorUserAgentIssue
            retVal = "NavigatorUserAgentIssue"
        Case IIC_WasmCrossOriginModuleSharingIssue
            retVal = "WasmCrossOriginModuleSharingIssue"
        Case IIC_GenericIssue
            retVal = "GenericIssue"
        Case Else
            Debug.Print "Warning, unknown value " & iic
    End Select
    InspectorIssueCodeToString = retVal
End Function

Public Function StringToInspectorIssueCode(ByVal s As String) As InspectorIssueCode
    Dim retVal As InspectorIssueCode
    Select Case s
        Case "SameSiteCookieIssue"
            retVal = IIC_SameSiteCookieIssue
        Case "MixedContentIssue"
            retVal = IIC_MixedContentIssue
        Case "BlockedByResponseIssue"
            retVal = IIC_BlockedByResponseIssue
        Case "HeavyAdIssue"
            retVal = IIC_HeavyAdIssue
        Case "ContentSecurityPolicyIssue"
            retVal = IIC_ContentSecurityPolicyIssue
        Case "SharedArrayBufferIssue"
            retVal = IIC_SharedArrayBufferIssue
        Case "TrustedWebActivityIssue"
            retVal = IIC_TrustedWebActivityIssue
        Case "LowTextContrastIssue"
            retVal = IIC_LowTextContrastIssue
        Case "CorsIssue"
            retVal = IIC_CorsIssue
        Case "AttributionReportingIssue"
            retVal = IIC_AttributionReportingIssue
        Case "QuirksModeIssue"
            retVal = IIC_QuirksModeIssue
        Case "NavigatorUserAgentIssue"
            retVal = IIC_NavigatorUserAgentIssue
        Case "WasmCrossOriginModuleSharingIssue"
            retVal = IIC_WasmCrossOriginModuleSharingIssue
        Case "GenericIssue"
            retVal = IIC_GenericIssue
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToInspectorIssueCode = retVal
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

' Returns the response body and size if it were re-encoded with the specified settings. Only
' applies to images.
' Returns:  - 
Public Function getEncodedResponse( _
    ByVal requestId AS string, _
    ByVal encoding AS string, _
    Optional ByVal quality AS Variant, _
    Optional ByVal sizeOnly AS Variant _
) AS Dictionary
    ' requestId: Network.RequestId Identifier of the network request to get content for.
    ' encoding: string The encoding to use.
    ' quality: number(optional) The quality of the encoding (0-1). (defaults to 1)
    ' sizeOnly: boolean(optional) Whether to only return the size information (defaults to false).

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)
    params("encoding") = CStr(encoding)
    If Not IsMissing(quality) Then params("quality") = CDbl(quality)
    If Not IsMissing(sizeOnly) Then params("sizeOnly") = CBool(sizeOnly)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Audits.getEncodedResponse", params)
    If cdp.ErrorCode = 0 Then
        Set getEncodedResponse = results
    End If
End Function

' Disables issues domain, prevents further issues from being reported to the client.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Audits.disable", params
End Sub

' Enables issues domain, sends the issues collected so far to the client by means of the
' `issueAdded` event.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Audits.enable", params
End Sub

' Runs the contrast check for the target page. Found issues are reported
' using Audits.issueAdded event.
Public Sub checkContrast( _
    Optional ByVal reportAAA AS Variant _
)
    ' reportAAA: boolean(optional) Whether to report WCAG AAA level issues. Default is false.

    Dim params As New Dictionary
    If Not IsMissing(reportAAA) Then params("reportAAA") = CBool(reportAAA)

    cdp.InvokeMethod "Audits.checkContrast", params
End Sub

