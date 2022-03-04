VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpEmulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpEmulation"
' Emulation
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Screen orientation.
' ScreenOrientation AS object

' 
' DisplayFeature AS object

' 
' MediaFeature AS object

' advance: If the scheduler runs out of immediate work, the virtual time base may fast forward to
'   allow the next delayed task (if any) to run; pause: The virtual time base may not advance;
'   pauseIfNetworkFetchesPending: The virtual time base may not advance if there are any pending
'   resource fetches.
' VirtualTimePolicy AS string
Public Enum VirtualTimePolicy
    VTP_advance
    VTP_pause
    VTP_pauseIfNetworkFetchesPending
End Enum

' Used to specify User Agent Cient Hints to emulate. See https://wicg.github.io/ua-client-hints
' UserAgentBrandVersion AS object

' Used to specify User Agent Cient Hints to emulate. See https://wicg.github.io/ua-client-hints
'   Missing optional values will be filled in by the target with what it would normally use.
' UserAgentMetadata AS object

' Enum of image types that can be disabled.
' DisabledImageType AS string
Public Enum DisabledImageType
    DIT_avif
    DIT_jxl
    DIT_webp
End Enum


Public Function VirtualTimePolicyToString(ByVal vtp As VirtualTimePolicy) As String
    Dim retVal As String
    Select Case vtp
        Case VTP_advance
            retVal = "advance"
        Case VTP_pause
            retVal = "pause"
        Case VTP_pauseIfNetworkFetchesPending
            retVal = "pauseIfNetworkFetchesPending"
        Case Else
            Debug.Print "Warning, unknown value " & vtp
    End Select
    VirtualTimePolicyToString = retVal
End Function

Public Function StringToVirtualTimePolicy(ByVal s As String) As VirtualTimePolicy
    Dim retVal As VirtualTimePolicy
    Select Case s
        Case "advance"
            retVal = VTP_advance
        Case "pause"
            retVal = VTP_pause
        Case "pauseIfNetworkFetchesPending"
            retVal = VTP_pauseIfNetworkFetchesPending
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToVirtualTimePolicy = retVal
End Function


Public Function DisabledImageTypeToString(ByVal dit As DisabledImageType) As String
    Dim retVal As String
    Select Case dit
        Case DIT_avif
            retVal = "avif"
        Case DIT_jxl
            retVal = "jxl"
        Case DIT_webp
            retVal = "webp"
        Case Else
            Debug.Print "Warning, unknown value " & dit
    End Select
    DisabledImageTypeToString = retVal
End Function

Public Function StringToDisabledImageType(ByVal s As String) As DisabledImageType
    Dim retVal As DisabledImageType
    Select Case s
        Case "avif"
            retVal = DIT_avif
        Case "jxl"
            retVal = DIT_jxl
        Case "webp"
            retVal = DIT_webp
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToDisabledImageType = retVal
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

' Tells whether emulation is supported.
' Returns: result - True if emulation is supported.
Public Function canEmulate() AS boolean
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Emulation.canEmulate", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then canEmulate = results("result")
    End If
End Function

' Clears the overridden device metrics.
Public Sub clearDeviceMetricsOverride()
    Dim params As New Dictionary
    cdp.InvokeMethod "Emulation.clearDeviceMetricsOverride", params
End Sub

' Clears the overridden Geolocation Position and Error.
Public Sub clearGeolocationOverride()
    Dim params As New Dictionary
    cdp.InvokeMethod "Emulation.clearGeolocationOverride", params
End Sub

' Requests that page scale factor is reset to initial values.
' Experimental
Public Sub resetPageScaleFactor()
    Dim params As New Dictionary
    cdp.InvokeMethod "Emulation.resetPageScaleFactor", params
End Sub

' Enables or disables simulating a focused and active page.
' Experimental
Public Sub setFocusEmulationEnabled( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean Whether to enable to disable focus emulation.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Emulation.setFocusEmulationEnabled", params
End Sub

' Automatically render all web contents using a dark theme.
' Experimental
Public Sub setAutoDarkModeOverride( _
    Optional ByVal enabled AS Variant _
)
    ' enabled: boolean(optional) Whether to enable or disable automatic dark mode.
    '   If not specified, any existing override will be cleared.

    Dim params As New Dictionary
    If Not IsMissing(enabled) Then params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Emulation.setAutoDarkModeOverride", params
End Sub

' Enables CPU throttling to emulate slow CPUs.
' Experimental
Public Sub setCPUThrottlingRate( _
    ByVal rate AS Double _
)
    ' rate: number Throttling rate as a slowdown factor (1 is no throttle, 2 is 2x slowdown, etc).

    Dim params As New Dictionary
    params("rate") = CDbl(rate)

    cdp.InvokeMethod "Emulation.setCPUThrottlingRate", params
End Sub

' Sets or clears an override of the default background color of the frame. This override is used
' if the content does not specify one.
Public Sub setDefaultBackgroundColorOverride( _
    Optional ByVal color AS Variant _
)
    ' color: DOM.RGBA(optional) RGBA of the default background color. If not specified, any existing override will be
    '   cleared.

    Dim params As New Dictionary
    If Not IsMissing(color) Then Set params("color") = color

    cdp.InvokeMethod "Emulation.setDefaultBackgroundColorOverride", params
End Sub

' Overrides the values of device screen dimensions (window.screen.width, window.screen.height,
' window.innerWidth, window.innerHeight, and "device-width"/"device-height"-related CSS media
' query results).
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
    Optional ByVal viewport AS Variant, _
    Optional ByVal displayFeature AS Variant _
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
    ' screenOrientation: ScreenOrientation(optional) Screen orientation override.
    ' viewport: Page.Viewport(optional) If set, the visible area of the page will be overridden to this viewport. This viewport
    '   change is not observed by the page, e.g. viewport-relative elements do not change positions.
    ' displayFeature: DisplayFeature(optional) If set, the display feature of a multi-segment screen. If not set, multi-segment support
    '   is turned-off.

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
    If Not IsMissing(displayFeature) Then Set params("displayFeature") = displayFeature

    cdp.InvokeMethod "Emulation.setDeviceMetricsOverride", params
End Sub

' 
' Experimental
Public Sub setScrollbarsHidden( _
    ByVal hidden AS boolean _
)
    ' hidden: boolean Whether scrollbars should be always hidden.

    Dim params As New Dictionary
    params("hidden") = CBool(hidden)

    cdp.InvokeMethod "Emulation.setScrollbarsHidden", params
End Sub

' 
' Experimental
Public Sub setDocumentCookieDisabled( _
    ByVal disabled AS boolean _
)
    ' disabled: boolean Whether document.coookie API should be disabled.

    Dim params As New Dictionary
    params("disabled") = CBool(disabled)

    cdp.InvokeMethod "Emulation.setDocumentCookieDisabled", params
End Sub

' 
' Experimental
Public Sub setEmitTouchEventsForMouse( _
    ByVal enabled AS boolean, _
    Optional ByVal configuration AS Variant _
)
    ' enabled: boolean Whether touch emulation based on mouse input should be enabled.
    ' configuration: string(optional) Touch/gesture events configuration. Default: current platform.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)
    If Not IsMissing(configuration) Then params("configuration") = CStr(configuration)

    cdp.InvokeMethod "Emulation.setEmitTouchEventsForMouse", params
End Sub

' Emulates the given media type or media feature for CSS media queries.
Public Sub setEmulatedMedia( _
    Optional ByVal media AS Variant, _
    Optional ByVal features AS Variant _
)
    ' media: string(optional) Media type to emulate. Empty string disables the override.
    ' features: array(optional) Media features to emulate.

    Dim params As New Dictionary
    If Not IsMissing(media) Then params("media") = CStr(media)
    If Not IsMissing(features) Then Set params("features") = features

    cdp.InvokeMethod "Emulation.setEmulatedMedia", params
End Sub

' Emulates the given vision deficiency.
' Experimental
Public Sub setEmulatedVisionDeficiency( _
    ByVal pType AS string _
)
    ' pType: string Vision deficiency to emulate.

    Dim params As New Dictionary
    params("type") = CStr(pType)

    cdp.InvokeMethod "Emulation.setEmulatedVisionDeficiency", params
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

    cdp.InvokeMethod "Emulation.setGeolocationOverride", params
End Sub

' Overrides the Idle state.
' Experimental
Public Sub setIdleOverride( _
    ByVal isUserActive AS boolean, _
    ByVal isScreenUnlocked AS boolean _
)
    ' isUserActive: boolean Mock isUserActive
    ' isScreenUnlocked: boolean Mock isScreenUnlocked

    Dim params As New Dictionary
    params("isUserActive") = CBool(isUserActive)
    params("isScreenUnlocked") = CBool(isScreenUnlocked)

    cdp.InvokeMethod "Emulation.setIdleOverride", params
End Sub

' Clears Idle state overrides.
' Experimental
Public Sub clearIdleOverride()
    Dim params As New Dictionary
    cdp.InvokeMethod "Emulation.clearIdleOverride", params
End Sub

' Overrides value returned by the javascript navigator object.
' Experimental
Public Sub setNavigatorOverrides( _
    ByVal platform AS string _
)
    ' platform: string The platform navigator.platform should return.

    Dim params As New Dictionary
    params("platform") = CStr(platform)

    cdp.InvokeMethod "Emulation.setNavigatorOverrides", params
End Sub

' Sets a specified page scale factor.
' Experimental
Public Sub setPageScaleFactor( _
    ByVal pageScaleFactor AS Double _
)
    ' pageScaleFactor: number Page scale factor.

    Dim params As New Dictionary
    params("pageScaleFactor") = CDbl(pageScaleFactor)

    cdp.InvokeMethod "Emulation.setPageScaleFactor", params
End Sub

' Switches script execution in the page.
Public Sub setScriptExecutionDisabled( _
    ByVal value AS boolean _
)
    ' value: boolean Whether script execution should be disabled in the page.

    Dim params As New Dictionary
    params("value") = CBool(value)

    cdp.InvokeMethod "Emulation.setScriptExecutionDisabled", params
End Sub

' Enables touch on platforms which do not support them.
Public Sub setTouchEmulationEnabled( _
    ByVal enabled AS boolean, _
    Optional ByVal maxTouchPoints AS Variant _
)
    ' enabled: boolean Whether the touch event emulation should be enabled.
    ' maxTouchPoints: integer(optional) Maximum touch points supported. Defaults to one.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)
    If Not IsMissing(maxTouchPoints) Then params("maxTouchPoints") = CLng(maxTouchPoints)

    cdp.InvokeMethod "Emulation.setTouchEmulationEnabled", params
End Sub

' Turns on virtual time for all frames (replacing real-time with a synthetic time source) and sets
' the current virtual time policy.  Note this supersedes any previous time budget.
' Returns: virtualTimeTicksBase - Absolute timestamp at which virtual time was first enabled (up time in milliseconds).
' Experimental
Public Function setVirtualTimePolicy( _
    ByVal policy AS string, _
    Optional ByVal budget AS Variant, _
    Optional ByVal maxVirtualTimeTaskStarvationCount AS Variant, _
    Optional ByVal waitForNavigation AS Variant, _
    Optional ByVal initialVirtualTime AS Variant _
) AS Double
    ' policy: VirtualTimePolicy 
    ' budget: number(optional) If set, after this many virtual milliseconds have elapsed virtual time will be paused and a
    '   virtualTimeBudgetExpired event is sent.
    ' maxVirtualTimeTaskStarvationCount: integer(optional) If set this specifies the maximum number of tasks that can be run before virtual is forced
    '   forwards to prevent deadlock.
    ' waitForNavigation: boolean(optional) If set the virtual time policy change should be deferred until any frame starts navigating.
    '   Note any previous deferred policy change is superseded.
    ' initialVirtualTime: Network.TimeSinceEpoch(optional) If set, base::Time::Now will be overridden to initially return this value.

    Dim params As New Dictionary
    params("policy") = CStr(policy)
    If Not IsMissing(budget) Then params("budget") = CDbl(budget)
    If Not IsMissing(maxVirtualTimeTaskStarvationCount) Then params("maxVirtualTimeTaskStarvationCount") = CLng(maxVirtualTimeTaskStarvationCount)
    If Not IsMissing(waitForNavigation) Then params("waitForNavigation") = CBool(waitForNavigation)
    If Not IsMissing(initialVirtualTime) Then params("initialVirtualTime") = CDbl(initialVirtualTime)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Emulation.setVirtualTimePolicy", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("virtualTimeTicksBase") Then setVirtualTimePolicy = results("virtualTimeTicksBase")
    End If
End Function

' Overrides default host system locale with the specified one.
' Experimental
Public Sub setLocaleOverride( _
    Optional ByVal locale AS Variant _
)
    ' locale: string(optional) ICU style C locale (e.g. "en_US"). If not specified or empty, disables the override and
    '   restores default host system locale.

    Dim params As New Dictionary
    If Not IsMissing(locale) Then params("locale") = CStr(locale)

    cdp.InvokeMethod "Emulation.setLocaleOverride", params
End Sub

' Overrides default host system timezone with the specified one.
' Experimental
Public Sub setTimezoneOverride( _
    ByVal timezoneId AS string _
)
    ' timezoneId: string The timezone identifier. If empty, disables the override and
    '   restores default host system timezone.

    Dim params As New Dictionary
    params("timezoneId") = CStr(timezoneId)

    cdp.InvokeMethod "Emulation.setTimezoneOverride", params
End Sub

' Resizes the frame/viewport of the page. Note that this does not affect the frame's container
' (e.g. browser window). Can be used to produce screenshots of the specified size. Not supported
' on Android.
' Experimental
Public Sub setVisibleSize( _
    ByVal width AS Long, _
    ByVal height AS Long _
)
    ' width: integer Frame width (DIP).
    ' height: integer Frame height (DIP).

    Dim params As New Dictionary
    params("width") = CLng(width)
    params("height") = CLng(height)

    cdp.InvokeMethod "Emulation.setVisibleSize", params
End Sub

' 
' Experimental
Public Sub setDisabledImageTypes( _
    ByVal imageTypes AS Collection _
)
    ' imageTypes: array Image types to disable.

    Dim params As New Dictionary
    Set params("imageTypes") = imageTypes

    cdp.InvokeMethod "Emulation.setDisabledImageTypes", params
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
    ' userAgentMetadata: UserAgentMetadata(optional) To be sent in Sec-CH-UA-* headers and returned in navigator.userAgentData

    Dim params As New Dictionary
    params("userAgent") = CStr(userAgent)
    If Not IsMissing(acceptLanguage) Then params("acceptLanguage") = CStr(acceptLanguage)
    If Not IsMissing(platform) Then params("platform") = CStr(platform)
    If Not IsMissing(userAgentMetadata) Then Set params("userAgentMetadata") = userAgentMetadata

    cdp.InvokeMethod "Emulation.setUserAgentOverride", params
End Sub

