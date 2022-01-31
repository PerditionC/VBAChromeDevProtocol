VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpHeadlessExperimental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpHeadlessExperimental"
' HeadlessExperimental [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Encoding options for a screenshot.
' ScreenshotParams AS object



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

' Sends a BeginFrame to the target and returns when the frame was completed. Optionally captures a
' screenshot from the resulting frame. Requires that the target was created with enabled
' BeginFrameControl. Designed for use with --run-all-compositor-stages-before-draw, see also
' https://goo.gl/3zHXhB for more background.
' Returns:  - 
Public Function beginFrame( _
    Optional ByVal frameTimeTicks AS Variant, _
    Optional ByVal interval AS Variant, _
    Optional ByVal noDisplayUpdates AS Variant, _
    Optional ByVal screenshot AS Variant _
) AS Dictionary
    ' frameTimeTicks: number(optional) Timestamp of this BeginFrame in Renderer TimeTicks (milliseconds of uptime). If not set,
    '   the current time will be used.
    ' interval: number(optional) The interval between BeginFrames that is reported to the compositor, in milliseconds.
    '   Defaults to a 60 frames/second interval, i.e. about 16.666 milliseconds.
    ' noDisplayUpdates: boolean(optional) Whether updates should not be committed and drawn onto the display. False by default. If
    '   true, only side effects of the BeginFrame will be run, such as layout and animations, but
    '   any visual updates may not be visible on the display or in screenshots.
    ' screenshot: ScreenshotParams(optional) If set, a screenshot of the frame will be captured and returned in the response. Otherwise,
    '   no screenshot will be captured. Note that capturing a screenshot can fail, for example,
    '   during renderer initialization. In such a case, no screenshot data will be returned.

    Dim params As New Dictionary
    If Not IsMissing(frameTimeTicks) Then params("frameTimeTicks") = CDbl(frameTimeTicks)
    If Not IsMissing(interval) Then params("interval") = CDbl(interval)
    If Not IsMissing(noDisplayUpdates) Then params("noDisplayUpdates") = CBool(noDisplayUpdates)
    If Not IsMissing(screenshot) Then Set params("screenshot") = screenshot

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("HeadlessExperimental.beginFrame", params)
    If cdp.ErrorCode = 0 Then
        Set beginFrame = results
    End If
End Function

' Disables headless events for the target.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "HeadlessExperimental.disable", params
End Sub

' Enables headless events for the target.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "HeadlessExperimental.enable", params
End Sub

