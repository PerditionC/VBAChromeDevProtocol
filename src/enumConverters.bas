Attribute VB_Name = "enumConverters"
' convert to/from enum/string
Option Explicit


' *** cdpBrowser

Public Function WindowStateToString(ByVal ws As windowState) As String
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

Public Function StringToWindowState(ByVal s As String) As windowState
    Dim retVal As windowState
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

