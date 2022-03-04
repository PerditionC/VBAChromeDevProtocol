VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpBrowser"
' Browser
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' 
' BrowserContextID AS string

' 
' WindowID AS integer

' The state of the browser window.
' WindowState AS string
Public Enum WindowState
    WS_normal
    WS_minimized
    WS_maximized
    WS_fullscreen
End Enum

' Browser window bounds information
' Bounds AS object

' 
' PermissionType AS string
Public Enum PermissionType
    PT_accessibilityEvents
    PT_audioCapture
    PT_backgroundSync
    PT_backgroundFetch
    PT_clipboardReadWrite
    PT_clipboardSanitizedWrite
    PT_displayCapture
    PT_durableStorage
    PT_flash
    PT_geolocation
    PT_midi
    PT_midiSysex
    PT_nfc
    PT_notifications
    PT_paymentHandler
    PT_periodicBackgroundSync
    PT_protectedMediaIdentifier
    PT_sensors
    PT_videoCapture
    PT_videoCapturePanTiltZoom
    PT_idleDetection
    PT_wakeLockScreen
    PT_wakeLockSystem
End Enum

' 
' PermissionSetting AS string
Public Enum PermissionSetting
    PS_granted
    PS_denied
    PS_prompt
End Enum

' Definition of PermissionDescriptor defined in the Permissions API:
'   https://w3c.github.io/permissions/#dictdef-permissiondescriptor.
' PermissionDescriptor AS object

' Browser command ids used by executeBrowserCommand.
' BrowserCommandId AS string
Public Enum BrowserCommandId
    BCI_openTabSearch
    BCI_closeTabSearch
    BCI_openTabPreview
End Enum

' Chrome histogram bucket.
' Bucket AS object

' Chrome histogram.
' Histogram AS object


Public Function WindowStateToString(ByVal ws As WindowState) As String
    Dim retVal As String
    Select Case ws
        Case WS_normal
            retVal = "normal"
        Case WS_minimized
            retVal = "minimized"
        Case WS_maximized
            retVal = "maximized"
        Case WS_fullscreen
            retVal = "fullscreen"
        Case Else
            Debug.Print "Warning, unknown value " & ws
    End Select
    WindowStateToString = retVal
End Function

Public Function StringToWindowState(ByVal s As String) As WindowState
    Dim retVal As WindowState
    Select Case s
        Case "normal"
            retVal = WS_normal
        Case "minimized"
            retVal = WS_minimized
        Case "maximized"
            retVal = WS_maximized
        Case "fullscreen"
            retVal = WS_fullscreen
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToWindowState = retVal
End Function


Public Function PermissionTypeToString(ByVal pt As PermissionType) As String
    Dim retVal As String
    Select Case pt
        Case PT_accessibilityEvents
            retVal = "accessibilityEvents"
        Case PT_audioCapture
            retVal = "audioCapture"
        Case PT_backgroundSync
            retVal = "backgroundSync"
        Case PT_backgroundFetch
            retVal = "backgroundFetch"
        Case PT_clipboardReadWrite
            retVal = "clipboardReadWrite"
        Case PT_clipboardSanitizedWrite
            retVal = "clipboardSanitizedWrite"
        Case PT_displayCapture
            retVal = "displayCapture"
        Case PT_durableStorage
            retVal = "durableStorage"
        Case PT_flash
            retVal = "flash"
        Case PT_geolocation
            retVal = "geolocation"
        Case PT_midi
            retVal = "midi"
        Case PT_midiSysex
            retVal = "midiSysex"
        Case PT_nfc
            retVal = "nfc"
        Case PT_notifications
            retVal = "notifications"
        Case PT_paymentHandler
            retVal = "paymentHandler"
        Case PT_periodicBackgroundSync
            retVal = "periodicBackgroundSync"
        Case PT_protectedMediaIdentifier
            retVal = "protectedMediaIdentifier"
        Case PT_sensors
            retVal = "sensors"
        Case PT_videoCapture
            retVal = "videoCapture"
        Case PT_videoCapturePanTiltZoom
            retVal = "videoCapturePanTiltZoom"
        Case PT_idleDetection
            retVal = "idleDetection"
        Case PT_wakeLockScreen
            retVal = "wakeLockScreen"
        Case PT_wakeLockSystem
            retVal = "wakeLockSystem"
        Case Else
            Debug.Print "Warning, unknown value " & pt
    End Select
    PermissionTypeToString = retVal
End Function

Public Function StringToPermissionType(ByVal s As String) As PermissionType
    Dim retVal As PermissionType
    Select Case s
        Case "accessibilityEvents"
            retVal = PT_accessibilityEvents
        Case "audioCapture"
            retVal = PT_audioCapture
        Case "backgroundSync"
            retVal = PT_backgroundSync
        Case "backgroundFetch"
            retVal = PT_backgroundFetch
        Case "clipboardReadWrite"
            retVal = PT_clipboardReadWrite
        Case "clipboardSanitizedWrite"
            retVal = PT_clipboardSanitizedWrite
        Case "displayCapture"
            retVal = PT_displayCapture
        Case "durableStorage"
            retVal = PT_durableStorage
        Case "flash"
            retVal = PT_flash
        Case "geolocation"
            retVal = PT_geolocation
        Case "midi"
            retVal = PT_midi
        Case "midiSysex"
            retVal = PT_midiSysex
        Case "nfc"
            retVal = PT_nfc
        Case "notifications"
            retVal = PT_notifications
        Case "paymentHandler"
            retVal = PT_paymentHandler
        Case "periodicBackgroundSync"
            retVal = PT_periodicBackgroundSync
        Case "protectedMediaIdentifier"
            retVal = PT_protectedMediaIdentifier
        Case "sensors"
            retVal = PT_sensors
        Case "videoCapture"
            retVal = PT_videoCapture
        Case "videoCapturePanTiltZoom"
            retVal = PT_videoCapturePanTiltZoom
        Case "idleDetection"
            retVal = PT_idleDetection
        Case "wakeLockScreen"
            retVal = PT_wakeLockScreen
        Case "wakeLockSystem"
            retVal = PT_wakeLockSystem
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToPermissionType = retVal
End Function


Public Function PermissionSettingToString(ByVal ps As PermissionSetting) As String
    Dim retVal As String
    Select Case ps
        Case PS_granted
            retVal = "granted"
        Case PS_denied
            retVal = "denied"
        Case PS_prompt
            retVal = "prompt"
        Case Else
            Debug.Print "Warning, unknown value " & ps
    End Select
    PermissionSettingToString = retVal
End Function

Public Function StringToPermissionSetting(ByVal s As String) As PermissionSetting
    Dim retVal As PermissionSetting
    Select Case s
        Case "granted"
            retVal = PS_granted
        Case "denied"
            retVal = PS_denied
        Case "prompt"
            retVal = PS_prompt
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToPermissionSetting = retVal
End Function


Public Function BrowserCommandIdToString(ByVal bci As BrowserCommandId) As String
    Dim retVal As String
    Select Case bci
        Case BCI_openTabSearch
            retVal = "openTabSearch"
        Case BCI_closeTabSearch
            retVal = "closeTabSearch"
        Case BCI_openTabPreview
            retVal = "openTabPreview"
        Case Else
            Debug.Print "Warning, unknown value " & bci
    End Select
    BrowserCommandIdToString = retVal
End Function

Public Function StringToBrowserCommandId(ByVal s As String) As BrowserCommandId
    Dim retVal As BrowserCommandId
    Select Case s
        Case "openTabSearch"
            retVal = BCI_openTabSearch
        Case "closeTabSearch"
            retVal = BCI_closeTabSearch
        Case "openTabPreview"
            retVal = BCI_openTabPreview
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToBrowserCommandId = retVal
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

' Set permission settings for given origin.
' Experimental
Public Sub setPermission( _
    ByVal permission AS Dictionary, _
    ByVal setting AS string, _
    Optional ByVal origin AS Variant, _
    Optional ByVal browserContextId AS Variant _
)
    ' permission: PermissionDescriptor Descriptor of permission to override.
    ' setting: PermissionSetting Setting of the permission.
    ' origin: string(optional) Origin the permission applies to, all origins if not specified.
    ' browserContextId: BrowserContextID(optional) Context to override. When omitted, default browser context is used.

    Dim params As New Dictionary
    Set params("permission") = permission
    params("setting") = CStr(setting)
    If Not IsMissing(origin) Then params("origin") = CStr(origin)
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)

    cdp.InvokeMethod "Browser.setPermission", params
End Sub

' Grant specific permissions to the given origin and reject all others.
' Experimental
Public Sub grantPermissions( _
    ByVal permissions AS Collection, _
    Optional ByVal origin AS Variant, _
    Optional ByVal browserContextId AS Variant _
)
    ' permissions: array 
    ' origin: string(optional) Origin the permission applies to, all origins if not specified.
    ' browserContextId: BrowserContextID(optional) BrowserContext to override permissions. When omitted, default browser context is used.

    Dim params As New Dictionary
    Set params("permissions") = permissions
    If Not IsMissing(origin) Then params("origin") = CStr(origin)
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)

    cdp.InvokeMethod "Browser.grantPermissions", params
End Sub

' Reset all permission management for all origins.
' Experimental
Public Sub resetPermissions( _
    Optional ByVal browserContextId AS Variant _
)
    ' browserContextId: BrowserContextID(optional) BrowserContext to reset permissions. When omitted, default browser context is used.

    Dim params As New Dictionary
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)

    cdp.InvokeMethod "Browser.resetPermissions", params
End Sub

' Set the behavior when downloading a file.
' Experimental
Public Sub setDownloadBehavior( _
    ByVal behavior AS string, _
    Optional ByVal browserContextId AS Variant, _
    Optional ByVal downloadPath AS Variant, _
    Optional ByVal eventsEnabled AS Variant _
)
    ' behavior: string Whether to allow all or deny all download requests, or use default Chrome behavior if
    '   available (otherwise deny). |allowAndName| allows download and names files according to
    '   their dowmload guids.
    ' browserContextId: BrowserContextID(optional) BrowserContext to set download behavior. When omitted, default browser context is used.
    ' downloadPath: string(optional) The default path to save downloaded files to. This is required if behavior is set to 'allow'
    '   or 'allowAndName'.
    ' eventsEnabled: boolean(optional) Whether to emit download events (defaults to false).

    Dim params As New Dictionary
    params("behavior") = CStr(behavior)
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)
    If Not IsMissing(downloadPath) Then params("downloadPath") = CStr(downloadPath)
    If Not IsMissing(eventsEnabled) Then params("eventsEnabled") = CBool(eventsEnabled)

    cdp.InvokeMethod "Browser.setDownloadBehavior", params
End Sub

' Cancel a download if in progress
' Experimental
Public Sub cancelDownload( _
    ByVal guid AS string, _
    Optional ByVal browserContextId AS Variant _
)
    ' guid: string Global unique identifier of the download.
    ' browserContextId: BrowserContextID(optional) BrowserContext to perform the action in. When omitted, default browser context is used.

    Dim params As New Dictionary
    params("guid") = CStr(guid)
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)

    cdp.InvokeMethod "Browser.cancelDownload", params
End Sub

' Close browser gracefully.
Public Sub closeBrowser()
    Dim params As New Dictionary
    cdp.InvokeMethod "Browser.close", params
End Sub

' Crashes browser on the main thread.
' Experimental
Public Sub crash()
    Dim params As New Dictionary
    cdp.InvokeMethod "Browser.crash", params
End Sub

' Crashes GPU process.
' Experimental
Public Sub crashGpuProcess()
    Dim params As New Dictionary
    cdp.InvokeMethod "Browser.crashGpuProcess", params
End Sub

' Returns version information.
' Returns:  - 
Public Function getVersion() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Browser.getVersion", params)
    If cdp.ErrorCode = 0 Then
        Set getVersion = results
    End If
End Function

' Returns the command line switches for the browser process if, and only if
' --enable-automation is on the commandline.
' Returns: arguments - Commandline parameters
' Experimental
Public Function getBrowserCommandLine() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Browser.getBrowserCommandLine", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("arguments") Then Set getBrowserCommandLine = results("arguments")
    End If
End Function

' Get Chrome histograms.
' Returns: histograms - Histograms.
' Experimental
Public Function getHistograms( _
    Optional ByVal query AS Variant, _
    Optional ByVal delta AS Variant _
) AS Collection
    ' query: string(optional) Requested substring in name. Only histograms which have query as a
    '   substring in their name are extracted. An empty or absent query returns
    '   all histograms.
    ' delta: boolean(optional) If true, retrieve delta since last call.

    Dim params As New Dictionary
    If Not IsMissing(query) Then params("query") = CStr(query)
    If Not IsMissing(delta) Then params("delta") = CBool(delta)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Browser.getHistograms", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("histograms") Then Set getHistograms = results("histograms")
    End If
End Function

' Get a Chrome histogram by name.
' Returns: histogram - Histogram.
' Experimental
Public Function getHistogram( _
    ByVal name AS string, _
    Optional ByVal delta AS Variant _
) AS Dictionary
    ' name: string Requested histogram name.
    ' delta: boolean(optional) If true, retrieve delta since last call.

    Dim params As New Dictionary
    params("name") = CStr(name)
    If Not IsMissing(delta) Then params("delta") = CBool(delta)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Browser.getHistogram", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("histogram") Then Set getHistogram = results("histogram")
    End If
End Function

' Get position and size of the browser window.
' Returns: bounds - Bounds information of the window. When window state is 'minimized', the restored window
 '   position and size are returned.
' Experimental
Public Function getWindowBounds( _
    ByVal windowId AS Long _
) AS Dictionary
    ' windowId: WindowID Browser window id.

    Dim params As New Dictionary
    params("windowId") = CLng(windowId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Browser.getWindowBounds", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("bounds") Then Set getWindowBounds = results("bounds")
    End If
End Function

' Get the browser window that contains the devtools target.
' Returns:  - 
' Experimental
Public Function getWindowForTarget( _
    Optional ByVal targetId AS Variant _
) AS Dictionary
    ' targetId: Target.TargetID(optional) Devtools agent host id. If called as a part of the session, associated targetId is used.

    Dim params As New Dictionary
    If Not IsMissing(targetId) Then params("targetId") = CStr(targetId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Browser.getWindowForTarget", params)
    If cdp.ErrorCode = 0 Then
        Set getWindowForTarget = results
    End If
End Function

' Set position and/or size of the browser window.
' Experimental
Public Sub setWindowBounds( _
    ByVal windowId AS Long, _
    ByVal bounds AS Dictionary _
)
    ' windowId: WindowID Browser window id.
    ' bounds: Bounds New window bounds. The 'minimized', 'maximized' and 'fullscreen' states cannot be combined
    '   with 'left', 'top', 'width' or 'height'. Leaves unspecified fields unchanged.

    Dim params As New Dictionary
    params("windowId") = CLng(windowId)
    Set params("bounds") = bounds

    cdp.InvokeMethod "Browser.setWindowBounds", params
End Sub

' Set dock tile details, platform-specific.
' Experimental
Public Sub setDockTile( _
    Optional ByVal badgeLabel AS Variant, _
    Optional ByVal image AS Variant _
)
    ' badgeLabel: string(optional) 
    ' image: binary(optional) Png encoded image.

    Dim params As New Dictionary
    If Not IsMissing(badgeLabel) Then params("badgeLabel") = CStr(badgeLabel)
    If Not IsMissing(image) Then params("image") = CStr(image)

    cdp.InvokeMethod "Browser.setDockTile", params
End Sub

' Invoke custom browser commands used by telemetry.
' Experimental
Public Sub executeBrowserCommand( _
    ByVal commandId AS string _
)
    ' commandId: BrowserCommandId 

    Dim params As New Dictionary
    params("commandId") = CStr(commandId)

    cdp.InvokeMethod "Browser.executeBrowserCommand", params
End Sub

