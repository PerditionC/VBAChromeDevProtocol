VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpPage"
' Page
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Unique frame identifier.
' FrameId AS string

' Indicates whether a frame has been identified as an ad.
' AdFrameType AS string
Public Enum AdFrameType
    AFT_none
    AFT_child
    AFT_root
End Enum

' 
' AdFrameExplanation AS string
Public Enum AdFrameExplanation
    AFE_ParentIsAd
    AFE_CreatedByAdScript
    AFE_MatchedBlockingRule
End Enum

' Indicates whether a frame has been identified as an ad and why.
' AdFrameStatus AS object

' Indicates whether the frame is a secure context and why it is the case.
' SecureContextType AS string
Public Enum SecureContextType
    SCT_Secure
    SCT_SecureLocalhost
    SCT_InsecureScheme
    SCT_InsecureAncestor
End Enum

' Indicates whether the frame is cross-origin isolated and why it is the case.
' CrossOriginIsolatedContextType AS string
Public Enum CrossOriginIsolatedContextType
    COICT_Isolated
    COICT_NotIsolated
    COICT_NotIsolatedFeatureDisabled
End Enum

' 
' GatedAPIFeatures AS string
Public Enum GatedAPIFeatures
    GAPIF_SharedArrayBuffers
    GAPIF_SharedArrayBuffersTransferAllowed
    GAPIF_PerformanceMeasureMemory
    GAPIF_PerformanceProfile
End Enum

' All Permissions Policy features. This enum should match the one defined
'   in third_party/blink/renderer/core/permissions_policy/permissions_policy_features.json5.
' PermissionsPolicyFeature AS string
Public Enum PermissionsPolicyFeature
    PPF_accelerometer
    PPF_ambientlightsensor
    PPF_attributionreporting
    PPF_autoplay
    PPF_camera
    PPF_chdpr
    PPF_chdevicememory
    PPF_chdownlink
    PPF_chect
    PPF_chpreferscolorscheme
    PPF_chrtt
    PPF_chua
    PPF_chuaarch
    PPF_chuabitness
    PPF_chuaplatform
    PPF_chuamodel
    PPF_chuamobile
    PPF_chuafullversion
    PPF_chuaplatformversion
    PPF_chuareduced
    PPF_chviewportheight
    PPF_chviewportwidth
    PPF_chwidth
    PPF_clipboardread
    PPF_clipboardwrite
    PPF_crossoriginisolated
    PPF_directsockets
    PPF_displaycapture
    PPF_documentdomain
    PPF_encryptedmedia
    PPF_executionwhileoutofviewport
    PPF_executionwhilenotrendered
    PPF_focuswithoutuseractivation
    PPF_fullscreen
    PPF_frobulate
    PPF_gamepad
    PPF_geolocation
    PPF_gyroscope
    PPF_hid
    PPF_idledetection
    PPF_interestcohort
    PPF_magnetometer
    PPF_microphone
    PPF_midi
    PPF_otpcredentials
    PPF_payment
    PPF_pictureinpicture
    PPF_publickeycredentialsget
    PPF_screenwakelock
    PPF_serial
    PPF_sharedautofill
    PPF_storageaccessapi
    PPF_syncxhr
    PPF_trusttokenredemption
    PPF_usb
    PPF_verticalscroll
    PPF_webshare
    PPF_windowplacement
    PPF_xrspatialtracking
End Enum

' Reason for a permissions policy feature to be disabled.
' PermissionsPolicyBlockReason AS string
Public Enum PermissionsPolicyBlockReason
    PPBR_Header
    PPBR_IframeAttribute
End Enum

' 
' PermissionsPolicyBlockLocator AS object

' 
' PermissionsPolicyFeatureState AS object

' Origin Trial(https://www.chromium.org/blink/origin-trials) support.
'   Status for an Origin Trial token.
' OriginTrialTokenStatus AS string
Public Enum OriginTrialTokenStatus
    OTTS_Success
    OTTS_NotSupported
    OTTS_Insecure
    OTTS_Expired
    OTTS_WrongOrigin
    OTTS_InvalidSignature
    OTTS_Malformed
    OTTS_WrongVersion
    OTTS_FeatureDisabled
    OTTS_TokenDisabled
    OTTS_FeatureDisabledForUser
    OTTS_UnknownTrial
End Enum

' Status for an Origin Trial.
' OriginTrialStatus AS string
Public Enum OriginTrialStatus
    OTS_Enabled
    OTS_ValidTokenNotProvided
    OTS_OSNotSupported
    OTS_TrialNotAllowed
End Enum

' 
' OriginTrialUsageRestriction AS string
Public Enum OriginTrialUsageRestriction
    OTUR_None
    OTUR_Subset
End Enum

' 
' OriginTrialToken AS object

' 
' OriginTrialTokenWithStatus AS object

' 
' OriginTrial AS object

' Information about the Frame on the page.
' Frame AS object

' Information about the Resource on the page.
' FrameResource AS object

' Information about the Frame hierarchy along with their cached resources.
' FrameResourceTree AS object

' Information about the Frame hierarchy.
' FrameTree AS object

' Unique script identifier.
' ScriptIdentifier AS string

' Transition type.
' TransitionType AS string
Public Enum TransitionType
    TT_link
    TT_typed
    TT_address_bar
    TT_auto_bookmark
    TT_auto_subframe
    TT_manual_subframe
    TT_generated
    TT_auto_toplevel
    TT_form_submit
    TT_reload
    TT_keyword
    TT_keyword_generated
    TT_other
End Enum

' Navigation history entry.
' NavigationEntry AS object

' Screencast frame metadata.
' ScreencastFrameMetadata AS object

' Javascript dialog type.
' DialogType AS string
Public Enum DialogType
    DT_alert
    DT_confirm
    DT_prompt
    DT_beforeunload
End Enum

' Error while paring app manifest.
' AppManifestError AS object

' Parsed app manifest properties.
' AppManifestParsedProperties AS object

' Layout viewport position and dimensions.
' LayoutViewport AS object

' Visual viewport position, dimensions, and scale.
' VisualViewport AS object

' Viewport for capturing screenshot.
' Viewport AS object

' Generic font families collection.
' FontFamilies AS object

' Default font sizes.
' FontSizes AS object

' 
' ClientNavigationReason AS string
Public Enum ClientNavigationReason
    CNR_formSubmissionGet
    CNR_formSubmissionPost
    CNR_httpHeaderRefresh
    CNR_scriptInitiated
    CNR_metaTagRefresh
    CNR_pageBlockInterstitial
    CNR_reload
    CNR_anchorClick
End Enum

' 
' ClientNavigationDisposition AS string
Public Enum ClientNavigationDisposition
    CND_currentTab
    CND_newTab
    CND_newWindow
    CND_download
End Enum

' 
' InstallabilityErrorArgument AS object

' The installability error
' InstallabilityError AS object

' The referring-policy used for the navigation.
' ReferrerPolicy AS string
Public Enum ReferrerPolicy
    RP_noReferrer
    RP_noReferrerWhenDowngrade
    RP_origin
    RP_originWhenCrossOrigin
    RP_sameOrigin
    RP_strictOrigin
    RP_strictOriginWhenCrossOrigin
    RP_unsafeUrl
End Enum

' Per-script compilation cache parameters for `Page.produceCompilationCache`
' CompilationCacheParams AS object

' The type of a frameNavigated event.
' NavigationType AS string
Public Enum NavigationType
    NT_Navigation
    NT_BackForwardCacheRestore
End Enum

' List of not restored reasons for back-forward cache.
' BackForwardCacheNotRestoredReason AS string
Public Enum BackForwardCacheNotRestoredReason
    BFCNRR_NotMainFrame
    BFCNRR_BackForwardCacheDisabled
    BFCNRR_RelatedActiveContentsExist
    BFCNRR_HTTPStatusNotOK
    BFCNRR_SchemeNotHTTPOrHTTPS
    BFCNRR_Loading
    BFCNRR_WasGrantedMediaAccess
    BFCNRR_DisableForRenderFrameHostCalled
    BFCNRR_DomainNotAllowed
    BFCNRR_HTTPMethodNotGET
    BFCNRR_SubframeIsNavigating
    BFCNRR_Timeout
    BFCNRR_CacheLimit
    BFCNRR_JavaScriptExecution
    BFCNRR_RendererProcessKilled
    BFCNRR_RendererProcessCrashed
    BFCNRR_GrantedMediaStreamAccess
    BFCNRR_SchedulerTrackedFeatureUsed
    BFCNRR_ConflictingBrowsingInstance
    BFCNRR_CacheFlushed
    BFCNRR_ServiceWorkerVersionActivation
    BFCNRR_SessionRestored
    BFCNRR_ServiceWorkerPostMessage
    BFCNRR_EnteredBackForwardCacheBeforeServiceWorkerHostAdded
    BFCNRR_RenderFrameHostReused_SameSite
    BFCNRR_RenderFrameHostReused_CrossSite
    BFCNRR_ServiceWorkerClaim
    BFCNRR_IgnoreEventAndEvict
    BFCNRR_HaveInnerContents
    BFCNRR_TimeoutPuttingInCache
    BFCNRR_BackForwardCacheDisabledByLowMemory
    BFCNRR_BackForwardCacheDisabledByCommandLine
    BFCNRR_NetworkRequestDatapipeDrainedAsBytesConsumer
    BFCNRR_NetworkRequestRedirected
    BFCNRR_NetworkRequestTimeout
    BFCNRR_NetworkExceedsBufferLimit
    BFCNRR_NavigationCancelledWhileRestoring
    BFCNRR_NotMostRecentNavigationEntry
    BFCNRR_BackForwardCacheDisabledForPrerender
    BFCNRR_UserAgentOverrideDiffers
    BFCNRR_ForegroundCacheLimit
    BFCNRR_BrowsingInstanceNotSwapped
    BFCNRR_BackForwardCacheDisabledForDelegate
    BFCNRR_OptInUnloadHeaderNotPresent
    BFCNRR_UnloadHandlerExistsInMainFrame
    BFCNRR_UnloadHandlerExistsInSubFrame
    BFCNRR_ServiceWorkerUnregistration
    BFCNRR_CacheControlNoStore
    BFCNRR_CacheControlNoStoreCookieModified
    BFCNRR_CacheControlNoStoreHTTPOnlyCookieModified
    BFCNRR_NoResponseHead
    BFCNRR_Unknown
    BFCNRR_ActivationNavigationsDisallowedForBug1234857
    BFCNRR_WebSocket
    BFCNRR_WebTransport
    BFCNRR_WebRTC
    BFCNRR_MainResourceHasCacheControlNoStore
    BFCNRR_MainResourceHasCacheControlNoCache
    BFCNRR_SubresourceHasCacheControlNoStore
    BFCNRR_SubresourceHasCacheControlNoCache
    BFCNRR_ContainsPlugins
    BFCNRR_DocumentLoaded
    BFCNRR_DedicatedWorkerOrWorklet
    BFCNRR_OutstandingNetworkRequestOthers
    BFCNRR_OutstandingIndexedDBTransaction
    BFCNRR_RequestedNotificationsPermission
    BFCNRR_RequestedMIDIPermission
    BFCNRR_RequestedAudioCapturePermission
    BFCNRR_RequestedVideoCapturePermission
    BFCNRR_RequestedBackForwardCacheBlockedSensors
    BFCNRR_RequestedBackgroundWorkPermission
    BFCNRR_BroadcastChannel
    BFCNRR_IndexedDBConnection
    BFCNRR_WebXR
    BFCNRR_SharedWorker
    BFCNRR_WebLocks
    BFCNRR_WebHID
    BFCNRR_WebShare
    BFCNRR_RequestedStorageAccessGrant
    BFCNRR_WebNfc
    BFCNRR_OutstandingNetworkRequestFetch
    BFCNRR_OutstandingNetworkRequestXHR
    BFCNRR_AppBanner
    BFCNRR_Printing
    BFCNRR_WebDatabase
    BFCNRR_PictureInPicture
    BFCNRR_Portal
    BFCNRR_SpeechRecognizer
    BFCNRR_IdleManager
    BFCNRR_PaymentManager
    BFCNRR_SpeechSynthesis
    BFCNRR_KeyboardLock
    BFCNRR_WebOTPService
    BFCNRR_OutstandingNetworkRequestDirectSocket
    BFCNRR_InjectedJavascript
    BFCNRR_InjectedStyleSheet
    BFCNRR_Dummy
    BFCNRR_ContentSecurityHandler
    BFCNRR_ContentWebAuthenticationAPI
    BFCNRR_ContentFileChooser
    BFCNRR_ContentSerial
    BFCNRR_ContentFileSystemAccess
    BFCNRR_ContentMediaDevicesDispatcherHost
    BFCNRR_ContentWebBluetooth
    BFCNRR_ContentWebUSB
    BFCNRR_ContentMediaSession
    BFCNRR_ContentMediaSessionService
    BFCNRR_ContentMediaPlay
    BFCNRR_EmbedderPopupBlockerTabHelper
    BFCNRR_EmbedderSafeBrowsingTriggeredPopupBlocker
    BFCNRR_EmbedderSafeBrowsingThreatDetails
    BFCNRR_EmbedderAppBannerManager
    BFCNRR_EmbedderDomDistillerViewerSource
    BFCNRR_EmbedderDomDistillerSelfDeletingRequestDelegate
    BFCNRR_EmbedderOomInterventionTabHelper
    BFCNRR_EmbedderOfflinePage
    BFCNRR_EmbedderChromePasswordManagerClientBindCredentialManager
    BFCNRR_EmbedderPermissionRequestManager
    BFCNRR_EmbedderModalDialog
    BFCNRR_EmbedderExtensions
    BFCNRR_EmbedderExtensionMessaging
    BFCNRR_EmbedderExtensionMessagingForOpenPort
    BFCNRR_EmbedderExtensionSentMessageToCachedFrame
End Enum

' Types of not restored reasons for back-forward cache.
' BackForwardCacheNotRestoredReasonType AS string
Public Enum BackForwardCacheNotRestoredReasonType
    BFCNRRT_SupportPending
    BFCNRRT_PageSupportNeeded
    BFCNRRT_Circumstantial
End Enum

' 
' BackForwardCacheNotRestoredExplanation AS object


Public Function AdFrameTypeToString(ByVal aft As AdFrameType) As String
    Dim retVal As String
    Select Case aft
        Case AFT_none
            retVal = "none"
        Case AFT_child
            retVal = "child"
        Case AFT_root
            retVal = "root"
        Case Else
            Debug.Print "Warning, unknown value " & aft
    End Select
    AdFrameTypeToString = retVal
End Function

Public Function StringToAdFrameType(ByVal s As String) As AdFrameType
    Dim retVal As AdFrameType
    Select Case s
        Case "none"
            retVal = AFT_none
        Case "child"
            retVal = AFT_child
        Case "root"
            retVal = AFT_root
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAdFrameType = retVal
End Function


Public Function AdFrameExplanationToString(ByVal afe As AdFrameExplanation) As String
    Dim retVal As String
    Select Case afe
        Case AFE_ParentIsAd
            retVal = "ParentIsAd"
        Case AFE_CreatedByAdScript
            retVal = "CreatedByAdScript"
        Case AFE_MatchedBlockingRule
            retVal = "MatchedBlockingRule"
        Case Else
            Debug.Print "Warning, unknown value " & afe
    End Select
    AdFrameExplanationToString = retVal
End Function

Public Function StringToAdFrameExplanation(ByVal s As String) As AdFrameExplanation
    Dim retVal As AdFrameExplanation
    Select Case s
        Case "ParentIsAd"
            retVal = AFE_ParentIsAd
        Case "CreatedByAdScript"
            retVal = AFE_CreatedByAdScript
        Case "MatchedBlockingRule"
            retVal = AFE_MatchedBlockingRule
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAdFrameExplanation = retVal
End Function


Public Function SecureContextTypeToString(ByVal sct As SecureContextType) As String
    Dim retVal As String
    Select Case sct
        Case SCT_Secure
            retVal = "Secure"
        Case SCT_SecureLocalhost
            retVal = "SecureLocalhost"
        Case SCT_InsecureScheme
            retVal = "InsecureScheme"
        Case SCT_InsecureAncestor
            retVal = "InsecureAncestor"
        Case Else
            Debug.Print "Warning, unknown value " & sct
    End Select
    SecureContextTypeToString = retVal
End Function

Public Function StringToSecureContextType(ByVal s As String) As SecureContextType
    Dim retVal As SecureContextType
    Select Case s
        Case "Secure"
            retVal = SCT_Secure
        Case "SecureLocalhost"
            retVal = SCT_SecureLocalhost
        Case "InsecureScheme"
            retVal = SCT_InsecureScheme
        Case "InsecureAncestor"
            retVal = SCT_InsecureAncestor
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSecureContextType = retVal
End Function


Public Function CrossOriginIsolatedContextTypeToString(ByVal coict As CrossOriginIsolatedContextType) As String
    Dim retVal As String
    Select Case coict
        Case COICT_Isolated
            retVal = "Isolated"
        Case COICT_NotIsolated
            retVal = "NotIsolated"
        Case COICT_NotIsolatedFeatureDisabled
            retVal = "NotIsolatedFeatureDisabled"
        Case Else
            Debug.Print "Warning, unknown value " & coict
    End Select
    CrossOriginIsolatedContextTypeToString = retVal
End Function

Public Function StringToCrossOriginIsolatedContextType(ByVal s As String) As CrossOriginIsolatedContextType
    Dim retVal As CrossOriginIsolatedContextType
    Select Case s
        Case "Isolated"
            retVal = COICT_Isolated
        Case "NotIsolated"
            retVal = COICT_NotIsolated
        Case "NotIsolatedFeatureDisabled"
            retVal = COICT_NotIsolatedFeatureDisabled
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToCrossOriginIsolatedContextType = retVal
End Function


Public Function GatedAPIFeaturesToString(ByVal gapif As GatedAPIFeatures) As String
    Dim retVal As String
    Select Case gapif
        Case GAPIF_SharedArrayBuffers
            retVal = "SharedArrayBuffers"
        Case GAPIF_SharedArrayBuffersTransferAllowed
            retVal = "SharedArrayBuffersTransferAllowed"
        Case GAPIF_PerformanceMeasureMemory
            retVal = "PerformanceMeasureMemory"
        Case GAPIF_PerformanceProfile
            retVal = "PerformanceProfile"
        Case Else
            Debug.Print "Warning, unknown value " & gapif
    End Select
    GatedAPIFeaturesToString = retVal
End Function

Public Function StringToGatedAPIFeatures(ByVal s As String) As GatedAPIFeatures
    Dim retVal As GatedAPIFeatures
    Select Case s
        Case "SharedArrayBuffers"
            retVal = GAPIF_SharedArrayBuffers
        Case "SharedArrayBuffersTransferAllowed"
            retVal = GAPIF_SharedArrayBuffersTransferAllowed
        Case "PerformanceMeasureMemory"
            retVal = GAPIF_PerformanceMeasureMemory
        Case "PerformanceProfile"
            retVal = GAPIF_PerformanceProfile
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToGatedAPIFeatures = retVal
End Function


Public Function PermissionsPolicyFeatureToString(ByVal ppf As PermissionsPolicyFeature) As String
    Dim retVal As String
    Select Case ppf
        Case PPF_accelerometer
            retVal = "accelerometer"
        Case PPF_ambientlightsensor
            retVal = "ambient-light-sensor"
        Case PPF_attributionreporting
            retVal = "attribution-reporting"
        Case PPF_autoplay
            retVal = "autoplay"
        Case PPF_camera
            retVal = "camera"
        Case PPF_chdpr
            retVal = "ch-dpr"
        Case PPF_chdevicememory
            retVal = "ch-device-memory"
        Case PPF_chdownlink
            retVal = "ch-downlink"
        Case PPF_chect
            retVal = "ch-ect"
        Case PPF_chpreferscolorscheme
            retVal = "ch-prefers-color-scheme"
        Case PPF_chrtt
            retVal = "ch-rtt"
        Case PPF_chua
            retVal = "ch-ua"
        Case PPF_chuaarch
            retVal = "ch-ua-arch"
        Case PPF_chuabitness
            retVal = "ch-ua-bitness"
        Case PPF_chuaplatform
            retVal = "ch-ua-platform"
        Case PPF_chuamodel
            retVal = "ch-ua-model"
        Case PPF_chuamobile
            retVal = "ch-ua-mobile"
        Case PPF_chuafullversion
            retVal = "ch-ua-full-version"
        Case PPF_chuaplatformversion
            retVal = "ch-ua-platform-version"
        Case PPF_chuareduced
            retVal = "ch-ua-reduced"
        Case PPF_chviewportheight
            retVal = "ch-viewport-height"
        Case PPF_chviewportwidth
            retVal = "ch-viewport-width"
        Case PPF_chwidth
            retVal = "ch-width"
        Case PPF_clipboardread
            retVal = "clipboard-read"
        Case PPF_clipboardwrite
            retVal = "clipboard-write"
        Case PPF_crossoriginisolated
            retVal = "cross-origin-isolated"
        Case PPF_directsockets
            retVal = "direct-sockets"
        Case PPF_displaycapture
            retVal = "display-capture"
        Case PPF_documentdomain
            retVal = "document-domain"
        Case PPF_encryptedmedia
            retVal = "encrypted-media"
        Case PPF_executionwhileoutofviewport
            retVal = "execution-while-out-of-viewport"
        Case PPF_executionwhilenotrendered
            retVal = "execution-while-not-rendered"
        Case PPF_focuswithoutuseractivation
            retVal = "focus-without-user-activation"
        Case PPF_fullscreen
            retVal = "fullscreen"
        Case PPF_frobulate
            retVal = "frobulate"
        Case PPF_gamepad
            retVal = "gamepad"
        Case PPF_geolocation
            retVal = "geolocation"
        Case PPF_gyroscope
            retVal = "gyroscope"
        Case PPF_hid
            retVal = "hid"
        Case PPF_idledetection
            retVal = "idle-detection"
        Case PPF_interestcohort
            retVal = "interest-cohort"
        Case PPF_magnetometer
            retVal = "magnetometer"
        Case PPF_microphone
            retVal = "microphone"
        Case PPF_midi
            retVal = "midi"
        Case PPF_otpcredentials
            retVal = "otp-credentials"
        Case PPF_payment
            retVal = "payment"
        Case PPF_pictureinpicture
            retVal = "picture-in-picture"
        Case PPF_publickeycredentialsget
            retVal = "publickey-credentials-get"
        Case PPF_screenwakelock
            retVal = "screen-wake-lock"
        Case PPF_serial
            retVal = "serial"
        Case PPF_sharedautofill
            retVal = "shared-autofill"
        Case PPF_storageaccessapi
            retVal = "storage-access-api"
        Case PPF_syncxhr
            retVal = "sync-xhr"
        Case PPF_trusttokenredemption
            retVal = "trust-token-redemption"
        Case PPF_usb
            retVal = "usb"
        Case PPF_verticalscroll
            retVal = "vertical-scroll"
        Case PPF_webshare
            retVal = "web-share"
        Case PPF_windowplacement
            retVal = "window-placement"
        Case PPF_xrspatialtracking
            retVal = "xr-spatial-tracking"
        Case Else
            Debug.Print "Warning, unknown value " & ppf
    End Select
    PermissionsPolicyFeatureToString = retVal
End Function

Public Function StringToPermissionsPolicyFeature(ByVal s As String) As PermissionsPolicyFeature
    Dim retVal As PermissionsPolicyFeature
    Select Case s
        Case "accelerometer"
            retVal = PPF_accelerometer
        Case "ambient-light-sensor"
            retVal = PPF_ambientlightsensor
        Case "attribution-reporting"
            retVal = PPF_attributionreporting
        Case "autoplay"
            retVal = PPF_autoplay
        Case "camera"
            retVal = PPF_camera
        Case "ch-dpr"
            retVal = PPF_chdpr
        Case "ch-device-memory"
            retVal = PPF_chdevicememory
        Case "ch-downlink"
            retVal = PPF_chdownlink
        Case "ch-ect"
            retVal = PPF_chect
        Case "ch-prefers-color-scheme"
            retVal = PPF_chpreferscolorscheme
        Case "ch-rtt"
            retVal = PPF_chrtt
        Case "ch-ua"
            retVal = PPF_chua
        Case "ch-ua-arch"
            retVal = PPF_chuaarch
        Case "ch-ua-bitness"
            retVal = PPF_chuabitness
        Case "ch-ua-platform"
            retVal = PPF_chuaplatform
        Case "ch-ua-model"
            retVal = PPF_chuamodel
        Case "ch-ua-mobile"
            retVal = PPF_chuamobile
        Case "ch-ua-full-version"
            retVal = PPF_chuafullversion
        Case "ch-ua-platform-version"
            retVal = PPF_chuaplatformversion
        Case "ch-ua-reduced"
            retVal = PPF_chuareduced
        Case "ch-viewport-height"
            retVal = PPF_chviewportheight
        Case "ch-viewport-width"
            retVal = PPF_chviewportwidth
        Case "ch-width"
            retVal = PPF_chwidth
        Case "clipboard-read"
            retVal = PPF_clipboardread
        Case "clipboard-write"
            retVal = PPF_clipboardwrite
        Case "cross-origin-isolated"
            retVal = PPF_crossoriginisolated
        Case "direct-sockets"
            retVal = PPF_directsockets
        Case "display-capture"
            retVal = PPF_displaycapture
        Case "document-domain"
            retVal = PPF_documentdomain
        Case "encrypted-media"
            retVal = PPF_encryptedmedia
        Case "execution-while-out-of-viewport"
            retVal = PPF_executionwhileoutofviewport
        Case "execution-while-not-rendered"
            retVal = PPF_executionwhilenotrendered
        Case "focus-without-user-activation"
            retVal = PPF_focuswithoutuseractivation
        Case "fullscreen"
            retVal = PPF_fullscreen
        Case "frobulate"
            retVal = PPF_frobulate
        Case "gamepad"
            retVal = PPF_gamepad
        Case "geolocation"
            retVal = PPF_geolocation
        Case "gyroscope"
            retVal = PPF_gyroscope
        Case "hid"
            retVal = PPF_hid
        Case "idle-detection"
            retVal = PPF_idledetection
        Case "interest-cohort"
            retVal = PPF_interestcohort
        Case "magnetometer"
            retVal = PPF_magnetometer
        Case "microphone"
            retVal = PPF_microphone
        Case "midi"
            retVal = PPF_midi
        Case "otp-credentials"
            retVal = PPF_otpcredentials
        Case "payment"
            retVal = PPF_payment
        Case "picture-in-picture"
            retVal = PPF_pictureinpicture
        Case "publickey-credentials-get"
            retVal = PPF_publickeycredentialsget
        Case "screen-wake-lock"
            retVal = PPF_screenwakelock
        Case "serial"
            retVal = PPF_serial
        Case "shared-autofill"
            retVal = PPF_sharedautofill
        Case "storage-access-api"
            retVal = PPF_storageaccessapi
        Case "sync-xhr"
            retVal = PPF_syncxhr
        Case "trust-token-redemption"
            retVal = PPF_trusttokenredemption
        Case "usb"
            retVal = PPF_usb
        Case "vertical-scroll"
            retVal = PPF_verticalscroll
        Case "web-share"
            retVal = PPF_webshare
        Case "window-placement"
            retVal = PPF_windowplacement
        Case "xr-spatial-tracking"
            retVal = PPF_xrspatialtracking
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToPermissionsPolicyFeature = retVal
End Function


Public Function PermissionsPolicyBlockReasonToString(ByVal ppbr As PermissionsPolicyBlockReason) As String
    Dim retVal As String
    Select Case ppbr
        Case PPBR_Header
            retVal = "Header"
        Case PPBR_IframeAttribute
            retVal = "IframeAttribute"
        Case Else
            Debug.Print "Warning, unknown value " & ppbr
    End Select
    PermissionsPolicyBlockReasonToString = retVal
End Function

Public Function StringToPermissionsPolicyBlockReason(ByVal s As String) As PermissionsPolicyBlockReason
    Dim retVal As PermissionsPolicyBlockReason
    Select Case s
        Case "Header"
            retVal = PPBR_Header
        Case "IframeAttribute"
            retVal = PPBR_IframeAttribute
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToPermissionsPolicyBlockReason = retVal
End Function


Public Function OriginTrialTokenStatusToString(ByVal otts As OriginTrialTokenStatus) As String
    Dim retVal As String
    Select Case otts
        Case OTTS_Success
            retVal = "Success"
        Case OTTS_NotSupported
            retVal = "NotSupported"
        Case OTTS_Insecure
            retVal = "Insecure"
        Case OTTS_Expired
            retVal = "Expired"
        Case OTTS_WrongOrigin
            retVal = "WrongOrigin"
        Case OTTS_InvalidSignature
            retVal = "InvalidSignature"
        Case OTTS_Malformed
            retVal = "Malformed"
        Case OTTS_WrongVersion
            retVal = "WrongVersion"
        Case OTTS_FeatureDisabled
            retVal = "FeatureDisabled"
        Case OTTS_TokenDisabled
            retVal = "TokenDisabled"
        Case OTTS_FeatureDisabledForUser
            retVal = "FeatureDisabledForUser"
        Case OTTS_UnknownTrial
            retVal = "UnknownTrial"
        Case Else
            Debug.Print "Warning, unknown value " & otts
    End Select
    OriginTrialTokenStatusToString = retVal
End Function

Public Function StringToOriginTrialTokenStatus(ByVal s As String) As OriginTrialTokenStatus
    Dim retVal As OriginTrialTokenStatus
    Select Case s
        Case "Success"
            retVal = OTTS_Success
        Case "NotSupported"
            retVal = OTTS_NotSupported
        Case "Insecure"
            retVal = OTTS_Insecure
        Case "Expired"
            retVal = OTTS_Expired
        Case "WrongOrigin"
            retVal = OTTS_WrongOrigin
        Case "InvalidSignature"
            retVal = OTTS_InvalidSignature
        Case "Malformed"
            retVal = OTTS_Malformed
        Case "WrongVersion"
            retVal = OTTS_WrongVersion
        Case "FeatureDisabled"
            retVal = OTTS_FeatureDisabled
        Case "TokenDisabled"
            retVal = OTTS_TokenDisabled
        Case "FeatureDisabledForUser"
            retVal = OTTS_FeatureDisabledForUser
        Case "UnknownTrial"
            retVal = OTTS_UnknownTrial
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToOriginTrialTokenStatus = retVal
End Function


Public Function OriginTrialStatusToString(ByVal ots As OriginTrialStatus) As String
    Dim retVal As String
    Select Case ots
        Case OTS_Enabled
            retVal = "Enabled"
        Case OTS_ValidTokenNotProvided
            retVal = "ValidTokenNotProvided"
        Case OTS_OSNotSupported
            retVal = "OSNotSupported"
        Case OTS_TrialNotAllowed
            retVal = "TrialNotAllowed"
        Case Else
            Debug.Print "Warning, unknown value " & ots
    End Select
    OriginTrialStatusToString = retVal
End Function

Public Function StringToOriginTrialStatus(ByVal s As String) As OriginTrialStatus
    Dim retVal As OriginTrialStatus
    Select Case s
        Case "Enabled"
            retVal = OTS_Enabled
        Case "ValidTokenNotProvided"
            retVal = OTS_ValidTokenNotProvided
        Case "OSNotSupported"
            retVal = OTS_OSNotSupported
        Case "TrialNotAllowed"
            retVal = OTS_TrialNotAllowed
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToOriginTrialStatus = retVal
End Function


Public Function OriginTrialUsageRestrictionToString(ByVal otur As OriginTrialUsageRestriction) As String
    Dim retVal As String
    Select Case otur
        Case OTUR_None
            retVal = "None"
        Case OTUR_Subset
            retVal = "Subset"
        Case Else
            Debug.Print "Warning, unknown value " & otur
    End Select
    OriginTrialUsageRestrictionToString = retVal
End Function

Public Function StringToOriginTrialUsageRestriction(ByVal s As String) As OriginTrialUsageRestriction
    Dim retVal As OriginTrialUsageRestriction
    Select Case s
        Case "None"
            retVal = OTUR_None
        Case "Subset"
            retVal = OTUR_Subset
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToOriginTrialUsageRestriction = retVal
End Function


Public Function TransitionTypeToString(ByVal tt As TransitionType) As String
    Dim retVal As String
    Select Case tt
        Case TT_link
            retVal = "link"
        Case TT_typed
            retVal = "typed"
        Case TT_address_bar
            retVal = "address_bar"
        Case TT_auto_bookmark
            retVal = "auto_bookmark"
        Case TT_auto_subframe
            retVal = "auto_subframe"
        Case TT_manual_subframe
            retVal = "manual_subframe"
        Case TT_generated
            retVal = "generated"
        Case TT_auto_toplevel
            retVal = "auto_toplevel"
        Case TT_form_submit
            retVal = "form_submit"
        Case TT_reload
            retVal = "reload"
        Case TT_keyword
            retVal = "keyword"
        Case TT_keyword_generated
            retVal = "keyword_generated"
        Case TT_other
            retVal = "other"
        Case Else
            Debug.Print "Warning, unknown value " & tt
    End Select
    TransitionTypeToString = retVal
End Function

Public Function StringToTransitionType(ByVal s As String) As TransitionType
    Dim retVal As TransitionType
    Select Case s
        Case "link"
            retVal = TT_link
        Case "typed"
            retVal = TT_typed
        Case "address_bar"
            retVal = TT_address_bar
        Case "auto_bookmark"
            retVal = TT_auto_bookmark
        Case "auto_subframe"
            retVal = TT_auto_subframe
        Case "manual_subframe"
            retVal = TT_manual_subframe
        Case "generated"
            retVal = TT_generated
        Case "auto_toplevel"
            retVal = TT_auto_toplevel
        Case "form_submit"
            retVal = TT_form_submit
        Case "reload"
            retVal = TT_reload
        Case "keyword"
            retVal = TT_keyword
        Case "keyword_generated"
            retVal = TT_keyword_generated
        Case "other"
            retVal = TT_other
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToTransitionType = retVal
End Function


Public Function DialogTypeToString(ByVal dt As DialogType) As String
    Dim retVal As String
    Select Case dt
        Case DT_alert
            retVal = "alert"
        Case DT_confirm
            retVal = "confirm"
        Case DT_prompt
            retVal = "prompt"
        Case DT_beforeunload
            retVal = "beforeunload"
        Case Else
            Debug.Print "Warning, unknown value " & dt
    End Select
    DialogTypeToString = retVal
End Function

Public Function StringToDialogType(ByVal s As String) As DialogType
    Dim retVal As DialogType
    Select Case s
        Case "alert"
            retVal = DT_alert
        Case "confirm"
            retVal = DT_confirm
        Case "prompt"
            retVal = DT_prompt
        Case "beforeunload"
            retVal = DT_beforeunload
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToDialogType = retVal
End Function


Public Function ClientNavigationReasonToString(ByVal cnr As ClientNavigationReason) As String
    Dim retVal As String
    Select Case cnr
        Case CNR_formSubmissionGet
            retVal = "formSubmissionGet"
        Case CNR_formSubmissionPost
            retVal = "formSubmissionPost"
        Case CNR_httpHeaderRefresh
            retVal = "httpHeaderRefresh"
        Case CNR_scriptInitiated
            retVal = "scriptInitiated"
        Case CNR_metaTagRefresh
            retVal = "metaTagRefresh"
        Case CNR_pageBlockInterstitial
            retVal = "pageBlockInterstitial"
        Case CNR_reload
            retVal = "reload"
        Case CNR_anchorClick
            retVal = "anchorClick"
        Case Else
            Debug.Print "Warning, unknown value " & cnr
    End Select
    ClientNavigationReasonToString = retVal
End Function

Public Function StringToClientNavigationReason(ByVal s As String) As ClientNavigationReason
    Dim retVal As ClientNavigationReason
    Select Case s
        Case "formSubmissionGet"
            retVal = CNR_formSubmissionGet
        Case "formSubmissionPost"
            retVal = CNR_formSubmissionPost
        Case "httpHeaderRefresh"
            retVal = CNR_httpHeaderRefresh
        Case "scriptInitiated"
            retVal = CNR_scriptInitiated
        Case "metaTagRefresh"
            retVal = CNR_metaTagRefresh
        Case "pageBlockInterstitial"
            retVal = CNR_pageBlockInterstitial
        Case "reload"
            retVal = CNR_reload
        Case "anchorClick"
            retVal = CNR_anchorClick
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToClientNavigationReason = retVal
End Function


Public Function ClientNavigationDispositionToString(ByVal cnd As ClientNavigationDisposition) As String
    Dim retVal As String
    Select Case cnd
        Case CND_currentTab
            retVal = "currentTab"
        Case CND_newTab
            retVal = "newTab"
        Case CND_newWindow
            retVal = "newWindow"
        Case CND_download
            retVal = "download"
        Case Else
            Debug.Print "Warning, unknown value " & cnd
    End Select
    ClientNavigationDispositionToString = retVal
End Function

Public Function StringToClientNavigationDisposition(ByVal s As String) As ClientNavigationDisposition
    Dim retVal As ClientNavigationDisposition
    Select Case s
        Case "currentTab"
            retVal = CND_currentTab
        Case "newTab"
            retVal = CND_newTab
        Case "newWindow"
            retVal = CND_newWindow
        Case "download"
            retVal = CND_download
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToClientNavigationDisposition = retVal
End Function


Public Function ReferrerPolicyToString(ByVal rp As ReferrerPolicy) As String
    Dim retVal As String
    Select Case rp
        Case RP_noReferrer
            retVal = "noReferrer"
        Case RP_noReferrerWhenDowngrade
            retVal = "noReferrerWhenDowngrade"
        Case RP_origin
            retVal = "origin"
        Case RP_originWhenCrossOrigin
            retVal = "originWhenCrossOrigin"
        Case RP_sameOrigin
            retVal = "sameOrigin"
        Case RP_strictOrigin
            retVal = "strictOrigin"
        Case RP_strictOriginWhenCrossOrigin
            retVal = "strictOriginWhenCrossOrigin"
        Case RP_unsafeUrl
            retVal = "unsafeUrl"
        Case Else
            Debug.Print "Warning, unknown value " & rp
    End Select
    ReferrerPolicyToString = retVal
End Function

Public Function StringToReferrerPolicy(ByVal s As String) As ReferrerPolicy
    Dim retVal As ReferrerPolicy
    Select Case s
        Case "noReferrer"
            retVal = RP_noReferrer
        Case "noReferrerWhenDowngrade"
            retVal = RP_noReferrerWhenDowngrade
        Case "origin"
            retVal = RP_origin
        Case "originWhenCrossOrigin"
            retVal = RP_originWhenCrossOrigin
        Case "sameOrigin"
            retVal = RP_sameOrigin
        Case "strictOrigin"
            retVal = RP_strictOrigin
        Case "strictOriginWhenCrossOrigin"
            retVal = RP_strictOriginWhenCrossOrigin
        Case "unsafeUrl"
            retVal = RP_unsafeUrl
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToReferrerPolicy = retVal
End Function


Public Function NavigationTypeToString(ByVal nt As NavigationType) As String
    Dim retVal As String
    Select Case nt
        Case NT_Navigation
            retVal = "Navigation"
        Case NT_BackForwardCacheRestore
            retVal = "BackForwardCacheRestore"
        Case Else
            Debug.Print "Warning, unknown value " & nt
    End Select
    NavigationTypeToString = retVal
End Function

Public Function StringToNavigationType(ByVal s As String) As NavigationType
    Dim retVal As NavigationType
    Select Case s
        Case "Navigation"
            retVal = NT_Navigation
        Case "BackForwardCacheRestore"
            retVal = NT_BackForwardCacheRestore
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToNavigationType = retVal
End Function


Public Function BackForwardCacheNotRestoredReasonToString(ByVal bfcnrr As BackForwardCacheNotRestoredReason) As String
    Dim retVal As String
    Select Case bfcnrr
        Case BFCNRR_NotMainFrame
            retVal = "NotMainFrame"
        Case BFCNRR_BackForwardCacheDisabled
            retVal = "BackForwardCacheDisabled"
        Case BFCNRR_RelatedActiveContentsExist
            retVal = "RelatedActiveContentsExist"
        Case BFCNRR_HTTPStatusNotOK
            retVal = "HTTPStatusNotOK"
        Case BFCNRR_SchemeNotHTTPOrHTTPS
            retVal = "SchemeNotHTTPOrHTTPS"
        Case BFCNRR_Loading
            retVal = "Loading"
        Case BFCNRR_WasGrantedMediaAccess
            retVal = "WasGrantedMediaAccess"
        Case BFCNRR_DisableForRenderFrameHostCalled
            retVal = "DisableForRenderFrameHostCalled"
        Case BFCNRR_DomainNotAllowed
            retVal = "DomainNotAllowed"
        Case BFCNRR_HTTPMethodNotGET
            retVal = "HTTPMethodNotGET"
        Case BFCNRR_SubframeIsNavigating
            retVal = "SubframeIsNavigating"
        Case BFCNRR_Timeout
            retVal = "Timeout"
        Case BFCNRR_CacheLimit
            retVal = "CacheLimit"
        Case BFCNRR_JavaScriptExecution
            retVal = "JavaScriptExecution"
        Case BFCNRR_RendererProcessKilled
            retVal = "RendererProcessKilled"
        Case BFCNRR_RendererProcessCrashed
            retVal = "RendererProcessCrashed"
        Case BFCNRR_GrantedMediaStreamAccess
            retVal = "GrantedMediaStreamAccess"
        Case BFCNRR_SchedulerTrackedFeatureUsed
            retVal = "SchedulerTrackedFeatureUsed"
        Case BFCNRR_ConflictingBrowsingInstance
            retVal = "ConflictingBrowsingInstance"
        Case BFCNRR_CacheFlushed
            retVal = "CacheFlushed"
        Case BFCNRR_ServiceWorkerVersionActivation
            retVal = "ServiceWorkerVersionActivation"
        Case BFCNRR_SessionRestored
            retVal = "SessionRestored"
        Case BFCNRR_ServiceWorkerPostMessage
            retVal = "ServiceWorkerPostMessage"
        Case BFCNRR_EnteredBackForwardCacheBeforeServiceWorkerHostAdded
            retVal = "EnteredBackForwardCacheBeforeServiceWorkerHostAdded"
        Case BFCNRR_RenderFrameHostReused_SameSite
            retVal = "RenderFrameHostReused_SameSite"
        Case BFCNRR_RenderFrameHostReused_CrossSite
            retVal = "RenderFrameHostReused_CrossSite"
        Case BFCNRR_ServiceWorkerClaim
            retVal = "ServiceWorkerClaim"
        Case BFCNRR_IgnoreEventAndEvict
            retVal = "IgnoreEventAndEvict"
        Case BFCNRR_HaveInnerContents
            retVal = "HaveInnerContents"
        Case BFCNRR_TimeoutPuttingInCache
            retVal = "TimeoutPuttingInCache"
        Case BFCNRR_BackForwardCacheDisabledByLowMemory
            retVal = "BackForwardCacheDisabledByLowMemory"
        Case BFCNRR_BackForwardCacheDisabledByCommandLine
            retVal = "BackForwardCacheDisabledByCommandLine"
        Case BFCNRR_NetworkRequestDatapipeDrainedAsBytesConsumer
            retVal = "NetworkRequestDatapipeDrainedAsBytesConsumer"
        Case BFCNRR_NetworkRequestRedirected
            retVal = "NetworkRequestRedirected"
        Case BFCNRR_NetworkRequestTimeout
            retVal = "NetworkRequestTimeout"
        Case BFCNRR_NetworkExceedsBufferLimit
            retVal = "NetworkExceedsBufferLimit"
        Case BFCNRR_NavigationCancelledWhileRestoring
            retVal = "NavigationCancelledWhileRestoring"
        Case BFCNRR_NotMostRecentNavigationEntry
            retVal = "NotMostRecentNavigationEntry"
        Case BFCNRR_BackForwardCacheDisabledForPrerender
            retVal = "BackForwardCacheDisabledForPrerender"
        Case BFCNRR_UserAgentOverrideDiffers
            retVal = "UserAgentOverrideDiffers"
        Case BFCNRR_ForegroundCacheLimit
            retVal = "ForegroundCacheLimit"
        Case BFCNRR_BrowsingInstanceNotSwapped
            retVal = "BrowsingInstanceNotSwapped"
        Case BFCNRR_BackForwardCacheDisabledForDelegate
            retVal = "BackForwardCacheDisabledForDelegate"
        Case BFCNRR_OptInUnloadHeaderNotPresent
            retVal = "OptInUnloadHeaderNotPresent"
        Case BFCNRR_UnloadHandlerExistsInMainFrame
            retVal = "UnloadHandlerExistsInMainFrame"
        Case BFCNRR_UnloadHandlerExistsInSubFrame
            retVal = "UnloadHandlerExistsInSubFrame"
        Case BFCNRR_ServiceWorkerUnregistration
            retVal = "ServiceWorkerUnregistration"
        Case BFCNRR_CacheControlNoStore
            retVal = "CacheControlNoStore"
        Case BFCNRR_CacheControlNoStoreCookieModified
            retVal = "CacheControlNoStoreCookieModified"
        Case BFCNRR_CacheControlNoStoreHTTPOnlyCookieModified
            retVal = "CacheControlNoStoreHTTPOnlyCookieModified"
        Case BFCNRR_NoResponseHead
            retVal = "NoResponseHead"
        Case BFCNRR_Unknown
            retVal = "Unknown"
        Case BFCNRR_ActivationNavigationsDisallowedForBug1234857
            retVal = "ActivationNavigationsDisallowedForBug1234857"
        Case BFCNRR_WebSocket
            retVal = "WebSocket"
        Case BFCNRR_WebTransport
            retVal = "WebTransport"
        Case BFCNRR_WebRTC
            retVal = "WebRTC"
        Case BFCNRR_MainResourceHasCacheControlNoStore
            retVal = "MainResourceHasCacheControlNoStore"
        Case BFCNRR_MainResourceHasCacheControlNoCache
            retVal = "MainResourceHasCacheControlNoCache"
        Case BFCNRR_SubresourceHasCacheControlNoStore
            retVal = "SubresourceHasCacheControlNoStore"
        Case BFCNRR_SubresourceHasCacheControlNoCache
            retVal = "SubresourceHasCacheControlNoCache"
        Case BFCNRR_ContainsPlugins
            retVal = "ContainsPlugins"
        Case BFCNRR_DocumentLoaded
            retVal = "DocumentLoaded"
        Case BFCNRR_DedicatedWorkerOrWorklet
            retVal = "DedicatedWorkerOrWorklet"
        Case BFCNRR_OutstandingNetworkRequestOthers
            retVal = "OutstandingNetworkRequestOthers"
        Case BFCNRR_OutstandingIndexedDBTransaction
            retVal = "OutstandingIndexedDBTransaction"
        Case BFCNRR_RequestedNotificationsPermission
            retVal = "RequestedNotificationsPermission"
        Case BFCNRR_RequestedMIDIPermission
            retVal = "RequestedMIDIPermission"
        Case BFCNRR_RequestedAudioCapturePermission
            retVal = "RequestedAudioCapturePermission"
        Case BFCNRR_RequestedVideoCapturePermission
            retVal = "RequestedVideoCapturePermission"
        Case BFCNRR_RequestedBackForwardCacheBlockedSensors
            retVal = "RequestedBackForwardCacheBlockedSensors"
        Case BFCNRR_RequestedBackgroundWorkPermission
            retVal = "RequestedBackgroundWorkPermission"
        Case BFCNRR_BroadcastChannel
            retVal = "BroadcastChannel"
        Case BFCNRR_IndexedDBConnection
            retVal = "IndexedDBConnection"
        Case BFCNRR_WebXR
            retVal = "WebXR"
        Case BFCNRR_SharedWorker
            retVal = "SharedWorker"
        Case BFCNRR_WebLocks
            retVal = "WebLocks"
        Case BFCNRR_WebHID
            retVal = "WebHID"
        Case BFCNRR_WebShare
            retVal = "WebShare"
        Case BFCNRR_RequestedStorageAccessGrant
            retVal = "RequestedStorageAccessGrant"
        Case BFCNRR_WebNfc
            retVal = "WebNfc"
        Case BFCNRR_OutstandingNetworkRequestFetch
            retVal = "OutstandingNetworkRequestFetch"
        Case BFCNRR_OutstandingNetworkRequestXHR
            retVal = "OutstandingNetworkRequestXHR"
        Case BFCNRR_AppBanner
            retVal = "AppBanner"
        Case BFCNRR_Printing
            retVal = "Printing"
        Case BFCNRR_WebDatabase
            retVal = "WebDatabase"
        Case BFCNRR_PictureInPicture
            retVal = "PictureInPicture"
        Case BFCNRR_Portal
            retVal = "Portal"
        Case BFCNRR_SpeechRecognizer
            retVal = "SpeechRecognizer"
        Case BFCNRR_IdleManager
            retVal = "IdleManager"
        Case BFCNRR_PaymentManager
            retVal = "PaymentManager"
        Case BFCNRR_SpeechSynthesis
            retVal = "SpeechSynthesis"
        Case BFCNRR_KeyboardLock
            retVal = "KeyboardLock"
        Case BFCNRR_WebOTPService
            retVal = "WebOTPService"
        Case BFCNRR_OutstandingNetworkRequestDirectSocket
            retVal = "OutstandingNetworkRequestDirectSocket"
        Case BFCNRR_InjectedJavascript
            retVal = "InjectedJavascript"
        Case BFCNRR_InjectedStyleSheet
            retVal = "InjectedStyleSheet"
        Case BFCNRR_Dummy
            retVal = "Dummy"
        Case BFCNRR_ContentSecurityHandler
            retVal = "ContentSecurityHandler"
        Case BFCNRR_ContentWebAuthenticationAPI
            retVal = "ContentWebAuthenticationAPI"
        Case BFCNRR_ContentFileChooser
            retVal = "ContentFileChooser"
        Case BFCNRR_ContentSerial
            retVal = "ContentSerial"
        Case BFCNRR_ContentFileSystemAccess
            retVal = "ContentFileSystemAccess"
        Case BFCNRR_ContentMediaDevicesDispatcherHost
            retVal = "ContentMediaDevicesDispatcherHost"
        Case BFCNRR_ContentWebBluetooth
            retVal = "ContentWebBluetooth"
        Case BFCNRR_ContentWebUSB
            retVal = "ContentWebUSB"
        Case BFCNRR_ContentMediaSession
            retVal = "ContentMediaSession"
        Case BFCNRR_ContentMediaSessionService
            retVal = "ContentMediaSessionService"
        Case BFCNRR_ContentMediaPlay
            retVal = "ContentMediaPlay"
        Case BFCNRR_EmbedderPopupBlockerTabHelper
            retVal = "EmbedderPopupBlockerTabHelper"
        Case BFCNRR_EmbedderSafeBrowsingTriggeredPopupBlocker
            retVal = "EmbedderSafeBrowsingTriggeredPopupBlocker"
        Case BFCNRR_EmbedderSafeBrowsingThreatDetails
            retVal = "EmbedderSafeBrowsingThreatDetails"
        Case BFCNRR_EmbedderAppBannerManager
            retVal = "EmbedderAppBannerManager"
        Case BFCNRR_EmbedderDomDistillerViewerSource
            retVal = "EmbedderDomDistillerViewerSource"
        Case BFCNRR_EmbedderDomDistillerSelfDeletingRequestDelegate
            retVal = "EmbedderDomDistillerSelfDeletingRequestDelegate"
        Case BFCNRR_EmbedderOomInterventionTabHelper
            retVal = "EmbedderOomInterventionTabHelper"
        Case BFCNRR_EmbedderOfflinePage
            retVal = "EmbedderOfflinePage"
        Case BFCNRR_EmbedderChromePasswordManagerClientBindCredentialManager
            retVal = "EmbedderChromePasswordManagerClientBindCredentialManager"
        Case BFCNRR_EmbedderPermissionRequestManager
            retVal = "EmbedderPermissionRequestManager"
        Case BFCNRR_EmbedderModalDialog
            retVal = "EmbedderModalDialog"
        Case BFCNRR_EmbedderExtensions
            retVal = "EmbedderExtensions"
        Case BFCNRR_EmbedderExtensionMessaging
            retVal = "EmbedderExtensionMessaging"
        Case BFCNRR_EmbedderExtensionMessagingForOpenPort
            retVal = "EmbedderExtensionMessagingForOpenPort"
        Case BFCNRR_EmbedderExtensionSentMessageToCachedFrame
            retVal = "EmbedderExtensionSentMessageToCachedFrame"
        Case Else
            Debug.Print "Warning, unknown value " & bfcnrr
    End Select
    BackForwardCacheNotRestoredReasonToString = retVal
End Function

Public Function StringToBackForwardCacheNotRestoredReason(ByVal s As String) As BackForwardCacheNotRestoredReason
    Dim retVal As BackForwardCacheNotRestoredReason
    Select Case s
        Case "NotMainFrame"
            retVal = BFCNRR_NotMainFrame
        Case "BackForwardCacheDisabled"
            retVal = BFCNRR_BackForwardCacheDisabled
        Case "RelatedActiveContentsExist"
            retVal = BFCNRR_RelatedActiveContentsExist
        Case "HTTPStatusNotOK"
            retVal = BFCNRR_HTTPStatusNotOK
        Case "SchemeNotHTTPOrHTTPS"
            retVal = BFCNRR_SchemeNotHTTPOrHTTPS
        Case "Loading"
            retVal = BFCNRR_Loading
        Case "WasGrantedMediaAccess"
            retVal = BFCNRR_WasGrantedMediaAccess
        Case "DisableForRenderFrameHostCalled"
            retVal = BFCNRR_DisableForRenderFrameHostCalled
        Case "DomainNotAllowed"
            retVal = BFCNRR_DomainNotAllowed
        Case "HTTPMethodNotGET"
            retVal = BFCNRR_HTTPMethodNotGET
        Case "SubframeIsNavigating"
            retVal = BFCNRR_SubframeIsNavigating
        Case "Timeout"
            retVal = BFCNRR_Timeout
        Case "CacheLimit"
            retVal = BFCNRR_CacheLimit
        Case "JavaScriptExecution"
            retVal = BFCNRR_JavaScriptExecution
        Case "RendererProcessKilled"
            retVal = BFCNRR_RendererProcessKilled
        Case "RendererProcessCrashed"
            retVal = BFCNRR_RendererProcessCrashed
        Case "GrantedMediaStreamAccess"
            retVal = BFCNRR_GrantedMediaStreamAccess
        Case "SchedulerTrackedFeatureUsed"
            retVal = BFCNRR_SchedulerTrackedFeatureUsed
        Case "ConflictingBrowsingInstance"
            retVal = BFCNRR_ConflictingBrowsingInstance
        Case "CacheFlushed"
            retVal = BFCNRR_CacheFlushed
        Case "ServiceWorkerVersionActivation"
            retVal = BFCNRR_ServiceWorkerVersionActivation
        Case "SessionRestored"
            retVal = BFCNRR_SessionRestored
        Case "ServiceWorkerPostMessage"
            retVal = BFCNRR_ServiceWorkerPostMessage
        Case "EnteredBackForwardCacheBeforeServiceWorkerHostAdded"
            retVal = BFCNRR_EnteredBackForwardCacheBeforeServiceWorkerHostAdded
        Case "RenderFrameHostReused_SameSite"
            retVal = BFCNRR_RenderFrameHostReused_SameSite
        Case "RenderFrameHostReused_CrossSite"
            retVal = BFCNRR_RenderFrameHostReused_CrossSite
        Case "ServiceWorkerClaim"
            retVal = BFCNRR_ServiceWorkerClaim
        Case "IgnoreEventAndEvict"
            retVal = BFCNRR_IgnoreEventAndEvict
        Case "HaveInnerContents"
            retVal = BFCNRR_HaveInnerContents
        Case "TimeoutPuttingInCache"
            retVal = BFCNRR_TimeoutPuttingInCache
        Case "BackForwardCacheDisabledByLowMemory"
            retVal = BFCNRR_BackForwardCacheDisabledByLowMemory
        Case "BackForwardCacheDisabledByCommandLine"
            retVal = BFCNRR_BackForwardCacheDisabledByCommandLine
        Case "NetworkRequestDatapipeDrainedAsBytesConsumer"
            retVal = BFCNRR_NetworkRequestDatapipeDrainedAsBytesConsumer
        Case "NetworkRequestRedirected"
            retVal = BFCNRR_NetworkRequestRedirected
        Case "NetworkRequestTimeout"
            retVal = BFCNRR_NetworkRequestTimeout
        Case "NetworkExceedsBufferLimit"
            retVal = BFCNRR_NetworkExceedsBufferLimit
        Case "NavigationCancelledWhileRestoring"
            retVal = BFCNRR_NavigationCancelledWhileRestoring
        Case "NotMostRecentNavigationEntry"
            retVal = BFCNRR_NotMostRecentNavigationEntry
        Case "BackForwardCacheDisabledForPrerender"
            retVal = BFCNRR_BackForwardCacheDisabledForPrerender
        Case "UserAgentOverrideDiffers"
            retVal = BFCNRR_UserAgentOverrideDiffers
        Case "ForegroundCacheLimit"
            retVal = BFCNRR_ForegroundCacheLimit
        Case "BrowsingInstanceNotSwapped"
            retVal = BFCNRR_BrowsingInstanceNotSwapped
        Case "BackForwardCacheDisabledForDelegate"
            retVal = BFCNRR_BackForwardCacheDisabledForDelegate
        Case "OptInUnloadHeaderNotPresent"
            retVal = BFCNRR_OptInUnloadHeaderNotPresent
        Case "UnloadHandlerExistsInMainFrame"
            retVal = BFCNRR_UnloadHandlerExistsInMainFrame
        Case "UnloadHandlerExistsInSubFrame"
            retVal = BFCNRR_UnloadHandlerExistsInSubFrame
        Case "ServiceWorkerUnregistration"
            retVal = BFCNRR_ServiceWorkerUnregistration
        Case "CacheControlNoStore"
            retVal = BFCNRR_CacheControlNoStore
        Case "CacheControlNoStoreCookieModified"
            retVal = BFCNRR_CacheControlNoStoreCookieModified
        Case "CacheControlNoStoreHTTPOnlyCookieModified"
            retVal = BFCNRR_CacheControlNoStoreHTTPOnlyCookieModified
        Case "NoResponseHead"
            retVal = BFCNRR_NoResponseHead
        Case "Unknown"
            retVal = BFCNRR_Unknown
        Case "ActivationNavigationsDisallowedForBug1234857"
            retVal = BFCNRR_ActivationNavigationsDisallowedForBug1234857
        Case "WebSocket"
            retVal = BFCNRR_WebSocket
        Case "WebTransport"
            retVal = BFCNRR_WebTransport
        Case "WebRTC"
            retVal = BFCNRR_WebRTC
        Case "MainResourceHasCacheControlNoStore"
            retVal = BFCNRR_MainResourceHasCacheControlNoStore
        Case "MainResourceHasCacheControlNoCache"
            retVal = BFCNRR_MainResourceHasCacheControlNoCache
        Case "SubresourceHasCacheControlNoStore"
            retVal = BFCNRR_SubresourceHasCacheControlNoStore
        Case "SubresourceHasCacheControlNoCache"
            retVal = BFCNRR_SubresourceHasCacheControlNoCache
        Case "ContainsPlugins"
            retVal = BFCNRR_ContainsPlugins
        Case "DocumentLoaded"
            retVal = BFCNRR_DocumentLoaded
        Case "DedicatedWorkerOrWorklet"
            retVal = BFCNRR_DedicatedWorkerOrWorklet
        Case "OutstandingNetworkRequestOthers"
            retVal = BFCNRR_OutstandingNetworkRequestOthers
        Case "OutstandingIndexedDBTransaction"
            retVal = BFCNRR_OutstandingIndexedDBTransaction
        Case "RequestedNotificationsPermission"
            retVal = BFCNRR_RequestedNotificationsPermission
        Case "RequestedMIDIPermission"
            retVal = BFCNRR_RequestedMIDIPermission
        Case "RequestedAudioCapturePermission"
            retVal = BFCNRR_RequestedAudioCapturePermission
        Case "RequestedVideoCapturePermission"
            retVal = BFCNRR_RequestedVideoCapturePermission
        Case "RequestedBackForwardCacheBlockedSensors"
            retVal = BFCNRR_RequestedBackForwardCacheBlockedSensors
        Case "RequestedBackgroundWorkPermission"
            retVal = BFCNRR_RequestedBackgroundWorkPermission
        Case "BroadcastChannel"
            retVal = BFCNRR_BroadcastChannel
        Case "IndexedDBConnection"
            retVal = BFCNRR_IndexedDBConnection
        Case "WebXR"
            retVal = BFCNRR_WebXR
        Case "SharedWorker"
            retVal = BFCNRR_SharedWorker
        Case "WebLocks"
            retVal = BFCNRR_WebLocks
        Case "WebHID"
            retVal = BFCNRR_WebHID
        Case "WebShare"
            retVal = BFCNRR_WebShare
        Case "RequestedStorageAccessGrant"
            retVal = BFCNRR_RequestedStorageAccessGrant
        Case "WebNfc"
            retVal = BFCNRR_WebNfc
        Case "OutstandingNetworkRequestFetch"
            retVal = BFCNRR_OutstandingNetworkRequestFetch
        Case "OutstandingNetworkRequestXHR"
            retVal = BFCNRR_OutstandingNetworkRequestXHR
        Case "AppBanner"
            retVal = BFCNRR_AppBanner
        Case "Printing"
            retVal = BFCNRR_Printing
        Case "WebDatabase"
            retVal = BFCNRR_WebDatabase
        Case "PictureInPicture"
            retVal = BFCNRR_PictureInPicture
        Case "Portal"
            retVal = BFCNRR_Portal
        Case "SpeechRecognizer"
            retVal = BFCNRR_SpeechRecognizer
        Case "IdleManager"
            retVal = BFCNRR_IdleManager
        Case "PaymentManager"
            retVal = BFCNRR_PaymentManager
        Case "SpeechSynthesis"
            retVal = BFCNRR_SpeechSynthesis
        Case "KeyboardLock"
            retVal = BFCNRR_KeyboardLock
        Case "WebOTPService"
            retVal = BFCNRR_WebOTPService
        Case "OutstandingNetworkRequestDirectSocket"
            retVal = BFCNRR_OutstandingNetworkRequestDirectSocket
        Case "InjectedJavascript"
            retVal = BFCNRR_InjectedJavascript
        Case "InjectedStyleSheet"
            retVal = BFCNRR_InjectedStyleSheet
        Case "Dummy"
            retVal = BFCNRR_Dummy
        Case "ContentSecurityHandler"
            retVal = BFCNRR_ContentSecurityHandler
        Case "ContentWebAuthenticationAPI"
            retVal = BFCNRR_ContentWebAuthenticationAPI
        Case "ContentFileChooser"
            retVal = BFCNRR_ContentFileChooser
        Case "ContentSerial"
            retVal = BFCNRR_ContentSerial
        Case "ContentFileSystemAccess"
            retVal = BFCNRR_ContentFileSystemAccess
        Case "ContentMediaDevicesDispatcherHost"
            retVal = BFCNRR_ContentMediaDevicesDispatcherHost
        Case "ContentWebBluetooth"
            retVal = BFCNRR_ContentWebBluetooth
        Case "ContentWebUSB"
            retVal = BFCNRR_ContentWebUSB
        Case "ContentMediaSession"
            retVal = BFCNRR_ContentMediaSession
        Case "ContentMediaSessionService"
            retVal = BFCNRR_ContentMediaSessionService
        Case "ContentMediaPlay"
            retVal = BFCNRR_ContentMediaPlay
        Case "EmbedderPopupBlockerTabHelper"
            retVal = BFCNRR_EmbedderPopupBlockerTabHelper
        Case "EmbedderSafeBrowsingTriggeredPopupBlocker"
            retVal = BFCNRR_EmbedderSafeBrowsingTriggeredPopupBlocker
        Case "EmbedderSafeBrowsingThreatDetails"
            retVal = BFCNRR_EmbedderSafeBrowsingThreatDetails
        Case "EmbedderAppBannerManager"
            retVal = BFCNRR_EmbedderAppBannerManager
        Case "EmbedderDomDistillerViewerSource"
            retVal = BFCNRR_EmbedderDomDistillerViewerSource
        Case "EmbedderDomDistillerSelfDeletingRequestDelegate"
            retVal = BFCNRR_EmbedderDomDistillerSelfDeletingRequestDelegate
        Case "EmbedderOomInterventionTabHelper"
            retVal = BFCNRR_EmbedderOomInterventionTabHelper
        Case "EmbedderOfflinePage"
            retVal = BFCNRR_EmbedderOfflinePage
        Case "EmbedderChromePasswordManagerClientBindCredentialManager"
            retVal = BFCNRR_EmbedderChromePasswordManagerClientBindCredentialManager
        Case "EmbedderPermissionRequestManager"
            retVal = BFCNRR_EmbedderPermissionRequestManager
        Case "EmbedderModalDialog"
            retVal = BFCNRR_EmbedderModalDialog
        Case "EmbedderExtensions"
            retVal = BFCNRR_EmbedderExtensions
        Case "EmbedderExtensionMessaging"
            retVal = BFCNRR_EmbedderExtensionMessaging
        Case "EmbedderExtensionMessagingForOpenPort"
            retVal = BFCNRR_EmbedderExtensionMessagingForOpenPort
        Case "EmbedderExtensionSentMessageToCachedFrame"
            retVal = BFCNRR_EmbedderExtensionSentMessageToCachedFrame
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToBackForwardCacheNotRestoredReason = retVal
End Function


Public Function BackForwardCacheNotRestoredReasonTypeToString(ByVal bfcnrrt As BackForwardCacheNotRestoredReasonType) As String
    Dim retVal As String
    Select Case bfcnrrt
        Case BFCNRRT_SupportPending
            retVal = "SupportPending"
        Case BFCNRRT_PageSupportNeeded
            retVal = "PageSupportNeeded"
        Case BFCNRRT_Circumstantial
            retVal = "Circumstantial"
        Case Else
            Debug.Print "Warning, unknown value " & bfcnrrt
    End Select
    BackForwardCacheNotRestoredReasonTypeToString = retVal
End Function

Public Function StringToBackForwardCacheNotRestoredReasonType(ByVal s As String) As BackForwardCacheNotRestoredReasonType
    Dim retVal As BackForwardCacheNotRestoredReasonType
    Select Case s
        Case "SupportPending"
            retVal = BFCNRRT_SupportPending
        Case "PageSupportNeeded"
            retVal = BFCNRRT_PageSupportNeeded
        Case "Circumstantial"
            retVal = BFCNRRT_Circumstantial
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToBackForwardCacheNotRestoredReasonType = retVal
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

' Deprecated, please use addScriptToEvaluateOnNewDocument instead.
' Returns: identifier - Identifier of the added script.
' Experimental
Public Function addScriptToEvaluateOnLoad( _
    ByVal scriptSource AS string _
) AS string
    ' scriptSource: string 

    Dim params As New Dictionary
    params("scriptSource") = CStr(scriptSource)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.addScriptToEvaluateOnLoad", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("identifier") Then addScriptToEvaluateOnLoad = results("identifier")
    End If
End Function

' Evaluates given script in every frame upon creation (before loading frame's scripts).
' Returns: identifier - Identifier of the added script.
Public Function addScriptToEvaluateOnNewDocument( _
    ByVal source AS string, _
    Optional ByVal worldName AS Variant, _
    Optional ByVal includeCommandLineAPI AS Variant _
) AS string
    ' source: string 
    ' worldName: string(optional) If specified, creates an isolated world with the given name and evaluates given script in it.
    '   This world name will be used as the ExecutionContextDescription::name when the corresponding
    '   event is emitted.
    ' includeCommandLineAPI: boolean(optional) Specifies whether command line API should be available to the script, defaults
    '   to false.

    Dim params As New Dictionary
    params("source") = CStr(source)
    If Not IsMissing(worldName) Then params("worldName") = CStr(worldName)
    If Not IsMissing(includeCommandLineAPI) Then params("includeCommandLineAPI") = CBool(includeCommandLineAPI)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.addScriptToEvaluateOnNewDocument", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("identifier") Then addScriptToEvaluateOnNewDocument = results("identifier")
    End If
End Function

' Brings page to front (activates tab).
Public Sub bringToFront()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.bringToFront", params
End Sub

' Capture page screenshot.
' Returns: data - Base64-encoded image data.
Public Function captureScreenshot( _
    Optional ByVal format AS Variant, _
    Optional ByVal quality AS Variant, _
    Optional ByVal clip AS Variant, _
    Optional ByVal fromSurface AS Variant, _
    Optional ByVal captureBeyondViewport AS Variant _
) AS String
    ' format: string(optional) Image compression format (defaults to png).
    ' quality: integer(optional) Compression quality from range [0..100] (jpeg only).
    ' clip: Viewport(optional) Capture the screenshot of a given region only.
    ' fromSurface: boolean(optional) Capture the screenshot from the surface, rather than the view. Defaults to true.
    ' captureBeyondViewport: boolean(optional) Capture the screenshot beyond the viewport. Defaults to false.

    Dim params As New Dictionary
    If Not IsMissing(format) Then params("format") = CStr(format)
    If Not IsMissing(quality) Then params("quality") = CLng(quality)
    If Not IsMissing(clip) Then Set params("clip") = clip
    If Not IsMissing(fromSurface) Then params("fromSurface") = CBool(fromSurface)
    If Not IsMissing(captureBeyondViewport) Then params("captureBeyondViewport") = CBool(captureBeyondViewport)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.captureScreenshot", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("data") Then captureScreenshot = results("data")
    End If
End Function

' Returns a snapshot of the page as a string. For MHTML format, the serialization includes
' iframes, shadow DOM, external resources, and element-inline styles.
' Returns: data - Serialized page data.
' Experimental
Public Function captureSnapshot( _
    Optional ByVal format AS Variant _
) AS string
    ' format: string(optional) Format (defaults to mhtml).

    Dim params As New Dictionary
    If Not IsMissing(format) Then params("format") = CStr(format)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.captureSnapshot", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("data") Then captureSnapshot = results("data")
    End If
End Function

' Clears the overridden device metrics.
' Experimental
Public Sub clearDeviceMetricsOverride()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.clearDeviceMetricsOverride", params
End Sub

' Clears the overridden Device Orientation.
' Experimental
Public Sub clearDeviceOrientationOverride()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.clearDeviceOrientationOverride", params
End Sub

' Clears the overridden Geolocation Position and Error.
Public Sub clearGeolocationOverride()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.clearGeolocationOverride", params
End Sub

' Creates an isolated world for the given frame.
' Returns: executionContextId - Execution context of the isolated world.
Public Function createIsolatedWorld( _
    ByVal frameId AS string, _
    Optional ByVal worldName AS Variant, _
    Optional ByVal grantUniveralAccess AS Variant _
) AS Long
    ' frameId: FrameId Id of the frame in which the isolated world should be created.
    ' worldName: string(optional) An optional name which is reported in the Execution Context.
    ' grantUniveralAccess: boolean(optional) Whether or not universal access should be granted to the isolated world. This is a powerful
    '   option, use with caution.

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)
    If Not IsMissing(worldName) Then params("worldName") = CStr(worldName)
    If Not IsMissing(grantUniveralAccess) Then params("grantUniveralAccess") = CBool(grantUniveralAccess)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.createIsolatedWorld", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("executionContextId") Then createIsolatedWorld = results("executionContextId")
    End If
End Function

' Deletes browser cookie with given name, domain and path.
' Experimental
Public Sub deleteCookie( _
    ByVal cookieName AS string, _
    ByVal url AS string _
)
    ' cookieName: string Name of the cookie to remove.
    ' url: string URL to match cooke domain and path.

    Dim params As New Dictionary
    params("cookieName") = CStr(cookieName)
    params("url") = CStr(url)

    cdp.InvokeMethod "Page.deleteCookie", params
End Sub

' Disables page domain notifications.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.disable", params
End Sub

' Enables page domain notifications.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.enable", params
End Sub

' 
' Returns:  - 
Public Function getAppManifest() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getAppManifest", params)
    If cdp.ErrorCode = 0 Then
        Set getAppManifest = results
    End If
End Function

' 
' Returns: installabilityErrors - 
' Experimental
Public Function getInstallabilityErrors() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getInstallabilityErrors", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("installabilityErrors") Then Set getInstallabilityErrors = results("installabilityErrors")
    End If
End Function

' 
' Returns: primaryIcon - 
' Experimental
Public Function getManifestIcons() AS String
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getManifestIcons", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("primaryIcon") Then getManifestIcons = results("primaryIcon")
    End If
End Function

' Returns the unique (PWA) app id.
' Only returns values if the feature flag 'WebAppEnableManifestId' is enabled
' Returns:  - 
' Experimental
Public Function getAppId() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getAppId", params)
    If cdp.ErrorCode = 0 Then
        Set getAppId = results
    End If
End Function

' Returns all browser cookies. Depending on the backend support, will return detailed cookie
' information in the `cookies` field.
' Returns: cookies - Array of cookie objects.
' Experimental
Public Function getCookies() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getCookies", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("cookies") Then Set getCookies = results("cookies")
    End If
End Function

' Returns present frame tree structure.
' Returns: frameTree - Present frame tree structure.
Public Function getFrameTree() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getFrameTree", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("frameTree") Then Set getFrameTree = results("frameTree")
    End If
End Function

' Returns metrics relating to the layouting of the page, such as viewport bounds/scale.
' Returns:  - 
Public Function getLayoutMetrics() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getLayoutMetrics", params)
    If cdp.ErrorCode = 0 Then
        Set getLayoutMetrics = results
    End If
End Function

' Returns navigation history for the current page.
' Returns:  - 
Public Function getNavigationHistory() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getNavigationHistory", params)
    If cdp.ErrorCode = 0 Then
        Set getNavigationHistory = results
    End If
End Function

' Resets navigation history for the current page.
Public Sub resetNavigationHistory()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.resetNavigationHistory", params
End Sub

' Returns content of the given resource.
' Returns:  - 
' Experimental
Public Function getResourceContent( _
    ByVal frameId AS string, _
    ByVal url AS string _
) AS Dictionary
    ' frameId: FrameId Frame id to get resource for.
    ' url: string URL of the resource to get content for.

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)
    params("url") = CStr(url)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getResourceContent", params)
    If cdp.ErrorCode = 0 Then
        Set getResourceContent = results
    End If
End Function

' Returns present frame / resource tree structure.
' Returns: frameTree - Present frame / resource tree structure.
' Experimental
Public Function getResourceTree() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getResourceTree", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("frameTree") Then Set getResourceTree = results("frameTree")
    End If
End Function

' Accepts or dismisses a JavaScript initiated dialog (alert, confirm, prompt, or onbeforeunload).
Public Sub handleJavaScriptDialog( _
    ByVal accept AS boolean, _
    Optional ByVal promptText AS Variant _
)
    ' accept: boolean Whether to accept or dismiss the dialog.
    ' promptText: string(optional) The text to enter into the dialog prompt before accepting. Used only if this is a prompt
    '   dialog.

    Dim params As New Dictionary
    params("accept") = CBool(accept)
    If Not IsMissing(promptText) Then params("promptText") = CStr(promptText)

    cdp.InvokeMethod "Page.handleJavaScriptDialog", params
End Sub

' Navigates current page to the given URL.
' Returns:  - 
Public Function navigate( _
    ByVal url AS string, _
    Optional ByVal referrer AS Variant, _
    Optional ByVal transitionType AS Variant, _
    Optional ByVal frameId AS Variant, _
    Optional ByVal referrerPolicy AS Variant _
) AS Dictionary
    ' url: string URL to navigate the page to.
    ' referrer: string(optional) Referrer URL.
    ' transitionType: TransitionType(optional) Intended transition type.
    ' frameId: FrameId(optional) Frame id to navigate, if not specified navigates the top frame.
    ' referrerPolicy: ReferrerPolicy(optional) Referrer-policy used for the navigation.

    Dim params As New Dictionary
    params("url") = CStr(url)
    If Not IsMissing(referrer) Then params("referrer") = CStr(referrer)
    If Not IsMissing(transitionType) Then params("transitionType") = CStr(transitionType)
    If Not IsMissing(frameId) Then params("frameId") = CStr(frameId)
    If Not IsMissing(referrerPolicy) Then params("referrerPolicy") = CStr(referrerPolicy)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.navigate", params)
    If cdp.ErrorCode = 0 Then
        Set navigate = results
    End If
End Function

' Navigates current page to the given history entry.
Public Sub navigateToHistoryEntry( _
    ByVal entryId AS Long _
)
    ' entryId: integer Unique id of the entry to navigate to.

    Dim params As New Dictionary
    params("entryId") = CLng(entryId)

    cdp.InvokeMethod "Page.navigateToHistoryEntry", params
End Sub

' Print page as PDF.
' Returns:  - 
Public Function printToPDF( _
    Optional ByVal landscape AS Variant, _
    Optional ByVal displayHeaderFooter AS Variant, _
    Optional ByVal printBackground AS Variant, _
    Optional ByVal pScale AS Variant, _
    Optional ByVal paperWidth AS Variant, _
    Optional ByVal paperHeight AS Variant, _
    Optional ByVal marginTop AS Variant, _
    Optional ByVal marginBottom AS Variant, _
    Optional ByVal marginLeft AS Variant, _
    Optional ByVal marginRight AS Variant, _
    Optional ByVal pageRanges AS Variant, _
    Optional ByVal ignoreInvalidPageRanges AS Variant, _
    Optional ByVal headerTemplate AS Variant, _
    Optional ByVal footerTemplate AS Variant, _
    Optional ByVal preferCSSPageSize AS Variant, _
    Optional ByVal transferMode AS Variant _
) AS Dictionary
    ' landscape: boolean(optional) Paper orientation. Defaults to false.
    ' displayHeaderFooter: boolean(optional) Display header and footer. Defaults to false.
    ' printBackground: boolean(optional) Print background graphics. Defaults to false.
    ' pScale: number(optional) Scale of the webpage rendering. Defaults to 1.
    ' paperWidth: number(optional) Paper width in inches. Defaults to 8.5 inches.
    ' paperHeight: number(optional) Paper height in inches. Defaults to 11 inches.
    ' marginTop: number(optional) Top margin in inches. Defaults to 1cm (~0.4 inches).
    ' marginBottom: number(optional) Bottom margin in inches. Defaults to 1cm (~0.4 inches).
    ' marginLeft: number(optional) Left margin in inches. Defaults to 1cm (~0.4 inches).
    ' marginRight: number(optional) Right margin in inches. Defaults to 1cm (~0.4 inches).
    ' pageRanges: string(optional) Paper ranges to print, e.g., '1-5, 8, 11-13'. Defaults to the empty string, which means
    '   print all pages.
    ' ignoreInvalidPageRanges: boolean(optional) Whether to silently ignore invalid but successfully parsed page ranges, such as '3-2'.
    '   Defaults to false.
    ' headerTemplate: string(optional) HTML template for the print header. Should be valid HTML markup with following
    '   classes used to inject printing values into them:
    '   - `date`: formatted print date
    '   - `title`: document title
    '   - `url`: document location
    '   - `pageNumber`: current page number
    '   - `totalPages`: total pages in the document
    '   
    '   For example, `<span class=title></span>` would generate span containing the title.
    ' footerTemplate: string(optional) HTML template for the print footer. Should use the same format as the `headerTemplate`.
    ' preferCSSPageSize: boolean(optional) Whether or not to prefer page size as defined by css. Defaults to false,
    '   in which case the content will be scaled to fit the paper size.
    ' transferMode: string(optional) return as stream

    Dim params As New Dictionary
    If Not IsMissing(landscape) Then params("landscape") = CBool(landscape)
    If Not IsMissing(displayHeaderFooter) Then params("displayHeaderFooter") = CBool(displayHeaderFooter)
    If Not IsMissing(printBackground) Then params("printBackground") = CBool(printBackground)
    If Not IsMissing(pScale) Then params("scale") = CDbl(pScale)
    If Not IsMissing(paperWidth) Then params("paperWidth") = CDbl(paperWidth)
    If Not IsMissing(paperHeight) Then params("paperHeight") = CDbl(paperHeight)
    If Not IsMissing(marginTop) Then params("marginTop") = CDbl(marginTop)
    If Not IsMissing(marginBottom) Then params("marginBottom") = CDbl(marginBottom)
    If Not IsMissing(marginLeft) Then params("marginLeft") = CDbl(marginLeft)
    If Not IsMissing(marginRight) Then params("marginRight") = CDbl(marginRight)
    If Not IsMissing(pageRanges) Then params("pageRanges") = CStr(pageRanges)
    If Not IsMissing(ignoreInvalidPageRanges) Then params("ignoreInvalidPageRanges") = CBool(ignoreInvalidPageRanges)
    If Not IsMissing(headerTemplate) Then params("headerTemplate") = CStr(headerTemplate)
    If Not IsMissing(footerTemplate) Then params("footerTemplate") = CStr(footerTemplate)
    If Not IsMissing(preferCSSPageSize) Then params("preferCSSPageSize") = CBool(preferCSSPageSize)
    If Not IsMissing(transferMode) Then params("transferMode") = CStr(transferMode)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.printToPDF", params)
    If cdp.ErrorCode = 0 Then
        Set printToPDF = results
    End If
End Function

' Reloads given page optionally ignoring the cache.
Public Sub reload( _
    Optional ByVal ignoreCache AS Variant, _
    Optional ByVal scriptToEvaluateOnLoad AS Variant _
)
    ' ignoreCache: boolean(optional) If true, browser cache is ignored (as if the user pressed Shift+refresh).
    ' scriptToEvaluateOnLoad: string(optional) If set, the script will be injected into all frames of the inspected page after reload.
    '   Argument will be ignored if reloading dataURL origin.

    Dim params As New Dictionary
    If Not IsMissing(ignoreCache) Then params("ignoreCache") = CBool(ignoreCache)
    If Not IsMissing(scriptToEvaluateOnLoad) Then params("scriptToEvaluateOnLoad") = CStr(scriptToEvaluateOnLoad)

    cdp.InvokeMethod "Page.reload", params
End Sub

' Deprecated, please use removeScriptToEvaluateOnNewDocument instead.
' Experimental
Public Sub removeScriptToEvaluateOnLoad( _
    ByVal identifier AS string _
)
    ' identifier: ScriptIdentifier 

    Dim params As New Dictionary
    params("identifier") = CStr(identifier)

    cdp.InvokeMethod "Page.removeScriptToEvaluateOnLoad", params
End Sub

' Removes given script from the list.
Public Sub removeScriptToEvaluateOnNewDocument( _
    ByVal identifier AS string _
)
    ' identifier: ScriptIdentifier 

    Dim params As New Dictionary
    params("identifier") = CStr(identifier)

    cdp.InvokeMethod "Page.removeScriptToEvaluateOnNewDocument", params
End Sub

' Acknowledges that a screencast frame has been received by the frontend.
' Experimental
Public Sub screencastFrameAck( _
    ByVal sessionId AS Long _
)
    ' sessionId: integer Frame number.

    Dim params As New Dictionary
    params("sessionId") = CLng(sessionId)

    cdp.InvokeMethod "Page.screencastFrameAck", params
End Sub

' Searches for given string in resource content.
' Returns: result - List of search matches.
' Experimental
Public Function searchInResource( _
    ByVal frameId AS string, _
    ByVal url AS string, _
    ByVal query AS string, _
    Optional ByVal caseSensitive AS Variant, _
    Optional ByVal isRegex AS Variant _
) AS Collection
    ' frameId: FrameId Frame id for resource to search in.
    ' url: string URL of the resource to search in.
    ' query: string String to search for.
    ' caseSensitive: boolean(optional) If true, search is case sensitive.
    ' isRegex: boolean(optional) If true, treats string parameter as regex.

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)
    params("url") = CStr(url)
    params("query") = CStr(query)
    If Not IsMissing(caseSensitive) Then params("caseSensitive") = CBool(caseSensitive)
    If Not IsMissing(isRegex) Then params("isRegex") = CBool(isRegex)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.searchInResource", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then Set searchInResource = results("result")
    End If
End Function

' Enable Chrome's experimental ad filter on all sites.
' Experimental
Public Sub setAdBlockingEnabled( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean Whether to block ads.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Page.setAdBlockingEnabled", params
End Sub

' Enable page Content Security Policy by-passing.
' Experimental
Public Sub setBypassCSP( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean Whether to bypass page CSP.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Page.setBypassCSP", params
End Sub

' Get Permissions Policy state on given frame.
' Returns: states - 
' Experimental
Public Function getPermissionsPolicyState( _
    ByVal frameId AS string _
) AS Collection
    ' frameId: FrameId 

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getPermissionsPolicyState", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("states") Then Set getPermissionsPolicyState = results("states")
    End If
End Function

' Get Origin Trials on given frame.
' Returns: originTrials - 
' Experimental
Public Function getOriginTrials( _
    ByVal frameId AS string _
) AS Collection
    ' frameId: FrameId 

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Page.getOriginTrials", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("originTrials") Then Set getOriginTrials = results("originTrials")
    End If
End Function

' Overrides the values of device screen dimensions (window.screen.width, window.screen.height,
' window.innerWidth, window.innerHeight, and "device-width"/"device-height"-related CSS media
' query results).
' Experimental
Public Sub setDeviceMetricsOverride( _
    ByVal width AS Long, _
    ByVal height AS Long, _
    ByVal deviceScaleFactor AS Double, _
    ByVal mobile AS boolean, _
    Optional ByVal pScale AS Variant, _
    Optional ByVal screenWidth AS Variant, _
    Optional ByVal screenHeight AS Variant, _
    Optional ByVal positionX AS Variant, _
    Optional ByVal positionY AS Variant, _
    Optional ByVal dontSetVisibleSize AS Variant, _
    Optional ByVal screenOrientation AS Variant, _
    Optional ByVal viewport AS Variant _
)
    ' width: integer Overriding width value in pixels (minimum 0, maximum 10000000). 0 disables the override.
    ' height: integer Overriding height value in pixels (minimum 0, maximum 10000000). 0 disables the override.
    ' deviceScaleFactor: number Overriding device scale factor value. 0 disables the override.
    ' mobile: boolean Whether to emulate mobile device. This includes viewport meta tag, overlay scrollbars, text
    '   autosizing and more.
    ' pScale: number(optional) Scale to apply to resulting view image.
    ' screenWidth: integer(optional) Overriding screen width value in pixels (minimum 0, maximum 10000000).
    ' screenHeight: integer(optional) Overriding screen height value in pixels (minimum 0, maximum 10000000).
    ' positionX: integer(optional) Overriding view X position on screen in pixels (minimum 0, maximum 10000000).
    ' positionY: integer(optional) Overriding view Y position on screen in pixels (minimum 0, maximum 10000000).
    ' dontSetVisibleSize: boolean(optional) Do not set visible view size, rely upon explicit setVisibleSize call.
    ' screenOrientation: Emulation.ScreenOrientation(optional) Screen orientation override.
    ' viewport: Viewport(optional) The viewport dimensions and scale. If not set, the override is cleared.

    Dim params As New Dictionary
    params("width") = CLng(width)
    params("height") = CLng(height)
    params("deviceScaleFactor") = CDbl(deviceScaleFactor)
    params("mobile") = CBool(mobile)
    If Not IsMissing(pScale) Then params("scale") = CDbl(pScale)
    If Not IsMissing(screenWidth) Then params("screenWidth") = CLng(screenWidth)
    If Not IsMissing(screenHeight) Then params("screenHeight") = CLng(screenHeight)
    If Not IsMissing(positionX) Then params("positionX") = CLng(positionX)
    If Not IsMissing(positionY) Then params("positionY") = CLng(positionY)
    If Not IsMissing(dontSetVisibleSize) Then params("dontSetVisibleSize") = CBool(dontSetVisibleSize)
    If Not IsMissing(screenOrientation) Then Set params("screenOrientation") = screenOrientation
    If Not IsMissing(viewport) Then Set params("viewport") = viewport

    cdp.InvokeMethod "Page.setDeviceMetricsOverride", params
End Sub

' Overrides the Device Orientation.
' Experimental
Public Sub setDeviceOrientationOverride( _
    ByVal alpha AS Double, _
    ByVal beta AS Double, _
    ByVal gamma AS Double _
)
    ' alpha: number Mock alpha
    ' beta: number Mock beta
    ' gamma: number Mock gamma

    Dim params As New Dictionary
    params("alpha") = CDbl(alpha)
    params("beta") = CDbl(beta)
    params("gamma") = CDbl(gamma)

    cdp.InvokeMethod "Page.setDeviceOrientationOverride", params
End Sub

' Set generic font families.
' Experimental
Public Sub setFontFamilies( _
    ByVal fontFamilies AS Dictionary _
)
    ' fontFamilies: FontFamilies Specifies font families to set. If a font family is not specified, it won't be changed.

    Dim params As New Dictionary
    Set params("fontFamilies") = fontFamilies

    cdp.InvokeMethod "Page.setFontFamilies", params
End Sub

' Set default font sizes.
' Experimental
Public Sub setFontSizes( _
    ByVal fontSizes AS Dictionary _
)
    ' fontSizes: FontSizes Specifies font sizes to set. If a font size is not specified, it won't be changed.

    Dim params As New Dictionary
    Set params("fontSizes") = fontSizes

    cdp.InvokeMethod "Page.setFontSizes", params
End Sub

' Sets given markup as the document's HTML.
Public Sub setDocumentContent( _
    ByVal frameId AS string, _
    ByVal html AS string _
)
    ' frameId: FrameId Frame id to set HTML for.
    ' html: string HTML content to set.

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)
    params("html") = CStr(html)

    cdp.InvokeMethod "Page.setDocumentContent", params
End Sub

' Set the behavior when downloading a file.
' Experimental
Public Sub setDownloadBehavior( _
    ByVal behavior AS string, _
    Optional ByVal downloadPath AS Variant _
)
    ' behavior: string Whether to allow all or deny all download requests, or use default Chrome behavior if
    '   available (otherwise deny).
    ' downloadPath: string(optional) The default path to save downloaded files to. This is required if behavior is set to 'allow'

    Dim params As New Dictionary
    params("behavior") = CStr(behavior)
    If Not IsMissing(downloadPath) Then params("downloadPath") = CStr(downloadPath)

    cdp.InvokeMethod "Page.setDownloadBehavior", params
End Sub

' Overrides the Geolocation Position or Error. Omitting any of the parameters emulates position
' unavailable.
Public Sub setGeolocationOverride( _
    Optional ByVal latitude AS Variant, _
    Optional ByVal longitude AS Variant, _
    Optional ByVal accuracy AS Variant _
)
    ' latitude: number(optional) Mock latitude
    ' longitude: number(optional) Mock longitude
    ' accuracy: number(optional) Mock accuracy

    Dim params As New Dictionary
    If Not IsMissing(latitude) Then params("latitude") = CDbl(latitude)
    If Not IsMissing(longitude) Then params("longitude") = CDbl(longitude)
    If Not IsMissing(accuracy) Then params("accuracy") = CDbl(accuracy)

    cdp.InvokeMethod "Page.setGeolocationOverride", params
End Sub

' Controls whether page will emit lifecycle events.
' Experimental
Public Sub setLifecycleEventsEnabled( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean If true, starts emitting lifecycle events.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Page.setLifecycleEventsEnabled", params
End Sub

' Toggles mouse event-based touch event emulation.
' Experimental
Public Sub setTouchEmulationEnabled( _
    ByVal enabled AS boolean, _
    Optional ByVal configuration AS Variant _
)
    ' enabled: boolean Whether the touch event emulation should be enabled.
    ' configuration: string(optional) Touch/gesture events configuration. Default: current platform.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)
    If Not IsMissing(configuration) Then params("configuration") = CStr(configuration)

    cdp.InvokeMethod "Page.setTouchEmulationEnabled", params
End Sub

' Starts sending each frame using the `screencastFrame` event.
' Experimental
Public Sub startScreencast( _
    Optional ByVal format AS Variant, _
    Optional ByVal quality AS Variant, _
    Optional ByVal maxWidth AS Variant, _
    Optional ByVal maxHeight AS Variant, _
    Optional ByVal everyNthFrame AS Variant _
)
    ' format: string(optional) Image compression format.
    ' quality: integer(optional) Compression quality from range [0..100].
    ' maxWidth: integer(optional) Maximum screenshot width.
    ' maxHeight: integer(optional) Maximum screenshot height.
    ' everyNthFrame: integer(optional) Send every n-th frame.

    Dim params As New Dictionary
    If Not IsMissing(format) Then params("format") = CStr(format)
    If Not IsMissing(quality) Then params("quality") = CLng(quality)
    If Not IsMissing(maxWidth) Then params("maxWidth") = CLng(maxWidth)
    If Not IsMissing(maxHeight) Then params("maxHeight") = CLng(maxHeight)
    If Not IsMissing(everyNthFrame) Then params("everyNthFrame") = CLng(everyNthFrame)

    cdp.InvokeMethod "Page.startScreencast", params
End Sub

' Force the page stop all navigations and pending resource fetches.
Public Sub stopLoading()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.stopLoading", params
End Sub

' Crashes renderer on the IO thread, generates minidumps.
' Experimental
Public Sub crash()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.crash", params
End Sub

' Tries to close page, running its beforeunload hooks, if any.
' Experimental
Public Sub closePage()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.close", params
End Sub

' Tries to update the web lifecycle state of the page.
' It will transition the page to the given state according to:
' https://github.com/WICG/web-lifecycle/
' Experimental
Public Sub setWebLifecycleState( _
    ByVal state AS string _
)
    ' state: string Target lifecycle state

    Dim params As New Dictionary
    params("state") = CStr(state)

    cdp.InvokeMethod "Page.setWebLifecycleState", params
End Sub

' Stops sending each frame in the `screencastFrame`.
' Experimental
Public Sub stopScreencast()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.stopScreencast", params
End Sub

' Forces compilation cache to be generated for every subresource script.
' See also: `Page.produceCompilationCache`.
' Experimental
Public Sub setProduceCompilationCache( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean 

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Page.setProduceCompilationCache", params
End Sub

' Requests backend to produce compilation cache for the specified scripts.
' Unlike setProduceCompilationCache, this allows client to only produce cache
' for specific scripts. `scripts` are appeneded to the list of scripts
' for which the cache for would produced. Disabling compilation cache with
' `setProduceCompilationCache` would reset all pending cache requests.
' The list may also be reset during page navigation.
' When script with a matching URL is encountered, the cache is optionally
' produced upon backend discretion, based on internal heuristics.
' See also: `Page.compilationCacheProduced`.
' Experimental
Public Sub produceCompilationCache( _
    ByVal scripts AS Collection _
)
    ' scripts: array 

    Dim params As New Dictionary
    Set params("scripts") = scripts

    cdp.InvokeMethod "Page.produceCompilationCache", params
End Sub

' Seeds compilation cache for given url. Compilation cache does not survive
' cross-process navigation.
' Experimental
Public Sub addCompilationCache( _
    ByVal url AS string, _
    ByVal data AS String _
)
    ' url: string 
    ' data: binary Base64-encoded data

    Dim params As New Dictionary
    params("url") = CStr(url)
    params("data") = CStr(data)

    cdp.InvokeMethod "Page.addCompilationCache", params
End Sub

' Clears seeded compilation cache.
' Experimental
Public Sub clearCompilationCache()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.clearCompilationCache", params
End Sub

' Generates a report for testing.
' Experimental
Public Sub generateTestReport( _
    ByVal message AS string, _
    Optional ByVal group AS Variant _
)
    ' message: string Message to be displayed in the report.
    ' group: string(optional) Specifies the endpoint group to deliver the report to.

    Dim params As New Dictionary
    params("message") = CStr(message)
    If Not IsMissing(group) Then params("group") = CStr(group)

    cdp.InvokeMethod "Page.generateTestReport", params
End Sub

' Pauses page execution. Can be resumed using generic Runtime.runIfWaitingForDebugger.
' Experimental
Public Sub waitForDebugger()
    Dim params As New Dictionary
    cdp.InvokeMethod "Page.waitForDebugger", params
End Sub

' Intercept file chooser requests and transfer control to protocol clients.
' When file chooser interception is enabled, native file chooser dialog is not shown.
' Instead, a protocol event `Page.fileChooserOpened` is emitted.
' Experimental
Public Sub setInterceptFileChooserDialog( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean 

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Page.setInterceptFileChooserDialog", params
End Sub

