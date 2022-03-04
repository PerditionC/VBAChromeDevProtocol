VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpInput"
' Input
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' 
' TouchPoint AS object

' 
' GestureSourceType AS string
Public Enum GestureSourceType
    GST_default
    GST_touch
    GST_mouse
End Enum

' 
' MouseButton AS string
Public Enum MouseButton
    MB_none
    MB_left
    MB_middle
    MB_right
    MB_back
    MB_forward
End Enum

' UTC time in seconds, counted from January 1, 1970.
' TimeSinceEpoch AS number

' 
' DragDataItem AS object

' 
' DragData AS object


Public Function GestureSourceTypeToString(ByVal gst As GestureSourceType) As String
    Dim retVal As String
    Select Case gst
        Case GST_default
            retVal = "default"
        Case GST_touch
            retVal = "touch"
        Case GST_mouse
            retVal = "mouse"
        Case Else
            Debug.Print "Warning, unknown value " & gst
    End Select
    GestureSourceTypeToString = retVal
End Function

Public Function StringToGestureSourceType(ByVal s As String) As GestureSourceType
    Dim retVal As GestureSourceType
    Select Case s
        Case "default"
            retVal = GST_default
        Case "touch"
            retVal = GST_touch
        Case "mouse"
            retVal = GST_mouse
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToGestureSourceType = retVal
End Function


Public Function MouseButtonToString(ByVal mb As MouseButton) As String
    Dim retVal As String
    Select Case mb
        Case MB_none
            retVal = "none"
        Case MB_left
            retVal = "left"
        Case MB_middle
            retVal = "middle"
        Case MB_right
            retVal = "right"
        Case MB_back
            retVal = "back"
        Case MB_forward
            retVal = "forward"
        Case Else
            Debug.Print "Warning, unknown value " & mb
    End Select
    MouseButtonToString = retVal
End Function

Public Function StringToMouseButton(ByVal s As String) As MouseButton
    Dim retVal As MouseButton
    Select Case s
        Case "none"
            retVal = MB_none
        Case "left"
            retVal = MB_left
        Case "middle"
            retVal = MB_middle
        Case "right"
            retVal = MB_right
        Case "back"
            retVal = MB_back
        Case "forward"
            retVal = MB_forward
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToMouseButton = retVal
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

' Dispatches a drag event into the page.
' Experimental
Public Sub dispatchDragEvent( _
    ByVal pType AS string, _
    ByVal x AS Double, _
    ByVal y AS Double, _
    ByVal data AS Dictionary, _
    Optional ByVal modifiers AS Variant _
)
    ' pType: string Type of the drag event.
    ' x: number X coordinate of the event relative to the main frame's viewport in CSS pixels.
    ' y: number Y coordinate of the event relative to the main frame's viewport in CSS pixels. 0 refers to
    '   the top of the viewport and Y increases as it proceeds towards the bottom of the viewport.
    ' data: DragData 
    ' modifiers: integer(optional) Bit field representing pressed modifier keys. Alt=1, Ctrl=2, Meta/Command=4, Shift=8
    '   (default: 0).

    Dim params As New Dictionary
    params("type") = CStr(pType)
    params("x") = CDbl(x)
    params("y") = CDbl(y)
    Set params("data") = data
    If Not IsMissing(modifiers) Then params("modifiers") = CLng(modifiers)

    cdp.InvokeMethod "Input.dispatchDragEvent", params
End Sub

' Dispatches a key event to the page.
Public Sub dispatchKeyEvent( _
    ByVal pType AS string, _
    Optional ByVal modifiers AS Variant, _
    Optional ByVal timestamp AS Variant, _
    Optional ByVal text AS Variant, _
    Optional ByVal unmodifiedText AS Variant, _
    Optional ByVal keyIdentifier AS Variant, _
    Optional ByVal code AS Variant, _
    Optional ByVal key AS Variant, _
    Optional ByVal windowsVirtualKeyCode AS Variant, _
    Optional ByVal nativeVirtualKeyCode AS Variant, _
    Optional ByVal autoRepeat AS Variant, _
    Optional ByVal isKeypad AS Variant, _
    Optional ByVal isSystemKey AS Variant, _
    Optional ByVal location AS Variant, _
    Optional ByVal commands AS Variant _
)
    ' pType: string Type of the key event.
    ' modifiers: integer(optional) Bit field representing pressed modifier keys. Alt=1, Ctrl=2, Meta/Command=4, Shift=8
    '   (default: 0).
    ' timestamp: TimeSinceEpoch(optional) Time at which the event occurred.
    ' text: string(optional) Text as generated by processing a virtual key code with a keyboard layout. Not needed for
    '   for `keyUp` and `rawKeyDown` events (default: "")
    ' unmodifiedText: string(optional) Text that would have been generated by the keyboard if no modifiers were pressed (except for
    '   shift). Useful for shortcut (accelerator) key handling (default: "").
    ' keyIdentifier: string(optional) Unique key identifier (e.g., 'U+0041') (default: "").
    ' code: string(optional) Unique DOM defined string value for each physical key (e.g., 'KeyA') (default: "").
    ' key: string(optional) Unique DOM defined string value describing the meaning of the key in the context of active
    '   modifiers, keyboard layout, etc (e.g., 'AltGr') (default: "").
    ' windowsVirtualKeyCode: integer(optional) Windows virtual key code (default: 0).
    ' nativeVirtualKeyCode: integer(optional) Native virtual key code (default: 0).
    ' autoRepeat: boolean(optional) Whether the event was generated from auto repeat (default: false).
    ' isKeypad: boolean(optional) Whether the event was generated from the keypad (default: false).
    ' isSystemKey: boolean(optional) Whether the event was a system key event (default: false).
    ' location: integer(optional) Whether the event was from the left or right side of the keyboard. 1=Left, 2=Right (default:
    '   0).
    ' commands: array(optional) Editing commands to send with the key event (e.g., 'selectAll') (default: []).
    '   These are related to but not equal the command names used in `document.execCommand` and NSStandardKeyBindingResponding.
    '   See https://source.chromium.org/chromium/chromium/src/+/master:third_party/blink/renderer/core/editing/commands/editor_command_names.h for valid command names.

    Dim params As New Dictionary
    params("type") = CStr(pType)
    If Not IsMissing(modifiers) Then params("modifiers") = CLng(modifiers)
    If Not IsMissing(timestamp) Then params("timestamp") = CDbl(timestamp)
    If Not IsMissing(text) Then params("text") = CStr(text)
    If Not IsMissing(unmodifiedText) Then params("unmodifiedText") = CStr(unmodifiedText)
    If Not IsMissing(keyIdentifier) Then params("keyIdentifier") = CStr(keyIdentifier)
    If Not IsMissing(code) Then params("code") = CStr(code)
    If Not IsMissing(key) Then params("key") = CStr(key)
    If Not IsMissing(windowsVirtualKeyCode) Then params("windowsVirtualKeyCode") = CLng(windowsVirtualKeyCode)
    If Not IsMissing(nativeVirtualKeyCode) Then params("nativeVirtualKeyCode") = CLng(nativeVirtualKeyCode)
    If Not IsMissing(autoRepeat) Then params("autoRepeat") = CBool(autoRepeat)
    If Not IsMissing(isKeypad) Then params("isKeypad") = CBool(isKeypad)
    If Not IsMissing(isSystemKey) Then params("isSystemKey") = CBool(isSystemKey)
    If Not IsMissing(location) Then params("location") = CLng(location)
    If Not IsMissing(commands) Then Set params("commands") = commands

    cdp.InvokeMethod "Input.dispatchKeyEvent", params
End Sub

' This method emulates inserting text that doesn't come from a key press,
' for example an emoji keyboard or an IME.
' Experimental
Public Sub insertText( _
    ByVal text AS string _
)
    ' text: string The text to insert.

    Dim params As New Dictionary
    params("text") = CStr(text)

    cdp.InvokeMethod "Input.insertText", params
End Sub

' This method sets the current candidate text for ime.
' Use imeCommitComposition to commit the final text.
' Use imeSetComposition with empty string as text to cancel composition.
' Experimental
Public Sub imeSetComposition( _
    ByVal text AS string, _
    ByVal selectionStart AS Long, _
    ByVal selectionEnd AS Long, _
    Optional ByVal replacementStart AS Variant, _
    Optional ByVal replacementEnd AS Variant _
)
    ' text: string The text to insert
    ' selectionStart: integer selection start
    ' selectionEnd: integer selection end
    ' replacementStart: integer(optional) replacement start
    ' replacementEnd: integer(optional) replacement end

    Dim params As New Dictionary
    params("text") = CStr(text)
    params("selectionStart") = CLng(selectionStart)
    params("selectionEnd") = CLng(selectionEnd)
    If Not IsMissing(replacementStart) Then params("replacementStart") = CLng(replacementStart)
    If Not IsMissing(replacementEnd) Then params("replacementEnd") = CLng(replacementEnd)

    cdp.InvokeMethod "Input.imeSetComposition", params
End Sub

' Dispatches a mouse event to the page.
Public Sub dispatchMouseEvent( _
    ByVal pType AS string, _
    ByVal x AS Double, _
    ByVal y AS Double, _
    Optional ByVal modifiers AS Variant, _
    Optional ByVal timestamp AS Variant, _
    Optional ByVal button AS Variant, _
    Optional ByVal buttons AS Variant, _
    Optional ByVal clickCount AS Variant, _
    Optional ByVal force AS Variant, _
    Optional ByVal tangentialPressure AS Variant, _
    Optional ByVal tiltX AS Variant, _
    Optional ByVal tiltY AS Variant, _
    Optional ByVal twist AS Variant, _
    Optional ByVal deltaX AS Variant, _
    Optional ByVal deltaY AS Variant, _
    Optional ByVal pointerType AS Variant _
)
    ' pType: string Type of the mouse event.
    ' x: number X coordinate of the event relative to the main frame's viewport in CSS pixels.
    ' y: number Y coordinate of the event relative to the main frame's viewport in CSS pixels. 0 refers to
    '   the top of the viewport and Y increases as it proceeds towards the bottom of the viewport.
    ' modifiers: integer(optional) Bit field representing pressed modifier keys. Alt=1, Ctrl=2, Meta/Command=4, Shift=8
    '   (default: 0).
    ' timestamp: TimeSinceEpoch(optional) Time at which the event occurred.
    ' button: MouseButton(optional) Mouse button (default: "none").
    ' buttons: integer(optional) A number indicating which buttons are pressed on the mouse when a mouse event is triggered.
    '   Left=1, Right=2, Middle=4, Back=8, Forward=16, None=0.
    ' clickCount: integer(optional) Number of times the mouse button was clicked (default: 0).
    ' force: number(optional) The normalized pressure, which has a range of [0,1] (default: 0).
    ' tangentialPressure: number(optional) The normalized tangential pressure, which has a range of [-1,1] (default: 0).
    ' tiltX: integer(optional) The plane angle between the Y-Z plane and the plane containing both the stylus axis and the Y axis, in degrees of the range [-90,90], a positive tiltX is to the right (default: 0).
    ' tiltY: integer(optional) The plane angle between the X-Z plane and the plane containing both the stylus axis and the X axis, in degrees of the range [-90,90], a positive tiltY is towards the user (default: 0).
    ' twist: integer(optional) The clockwise rotation of a pen stylus around its own major axis, in degrees in the range [0,359] (default: 0).
    ' deltaX: number(optional) X delta in CSS pixels for mouse wheel event (default: 0).
    ' deltaY: number(optional) Y delta in CSS pixels for mouse wheel event (default: 0).
    ' pointerType: string(optional) Pointer type (default: "mouse").

    Dim params As New Dictionary
    params("type") = CStr(pType)
    params("x") = CDbl(x)
    params("y") = CDbl(y)
    If Not IsMissing(modifiers) Then params("modifiers") = CLng(modifiers)
    If Not IsMissing(timestamp) Then params("timestamp") = CDbl(timestamp)
    If Not IsMissing(button) Then params("button") = CStr(button)
    If Not IsMissing(buttons) Then params("buttons") = CLng(buttons)
    If Not IsMissing(clickCount) Then params("clickCount") = CLng(clickCount)
    If Not IsMissing(force) Then params("force") = CDbl(force)
    If Not IsMissing(tangentialPressure) Then params("tangentialPressure") = CDbl(tangentialPressure)
    If Not IsMissing(tiltX) Then params("tiltX") = CLng(tiltX)
    If Not IsMissing(tiltY) Then params("tiltY") = CLng(tiltY)
    If Not IsMissing(twist) Then params("twist") = CLng(twist)
    If Not IsMissing(deltaX) Then params("deltaX") = CDbl(deltaX)
    If Not IsMissing(deltaY) Then params("deltaY") = CDbl(deltaY)
    If Not IsMissing(pointerType) Then params("pointerType") = CStr(pointerType)

    cdp.InvokeMethod "Input.dispatchMouseEvent", params
End Sub

' Dispatches a touch event to the page.
Public Sub dispatchTouchEvent( _
    ByVal pType AS string, _
    ByVal touchPoints AS Collection, _
    Optional ByVal modifiers AS Variant, _
    Optional ByVal timestamp AS Variant _
)
    ' pType: string Type of the touch event. TouchEnd and TouchCancel must not contain any touch points, while
    '   TouchStart and TouchMove must contains at least one.
    ' touchPoints: array Active touch points on the touch device. One event per any changed point (compared to
    '   previous touch event in a sequence) is generated, emulating pressing/moving/releasing points
    '   one by one.
    ' modifiers: integer(optional) Bit field representing pressed modifier keys. Alt=1, Ctrl=2, Meta/Command=4, Shift=8
    '   (default: 0).
    ' timestamp: TimeSinceEpoch(optional) Time at which the event occurred.

    Dim params As New Dictionary
    params("type") = CStr(pType)
    Set params("touchPoints") = touchPoints
    If Not IsMissing(modifiers) Then params("modifiers") = CLng(modifiers)
    If Not IsMissing(timestamp) Then params("timestamp") = CDbl(timestamp)

    cdp.InvokeMethod "Input.dispatchTouchEvent", params
End Sub

' Emulates touch event from the mouse event parameters.
' Experimental
Public Sub emulateTouchFromMouseEvent( _
    ByVal pType AS string, _
    ByVal x AS Long, _
    ByVal y AS Long, _
    ByVal button AS string, _
    Optional ByVal timestamp AS Variant, _
    Optional ByVal deltaX AS Variant, _
    Optional ByVal deltaY AS Variant, _
    Optional ByVal modifiers AS Variant, _
    Optional ByVal clickCount AS Variant _
)
    ' pType: string Type of the mouse event.
    ' x: integer X coordinate of the mouse pointer in DIP.
    ' y: integer Y coordinate of the mouse pointer in DIP.
    ' button: MouseButton Mouse button. Only "none", "left", "right" are supported.
    ' timestamp: TimeSinceEpoch(optional) Time at which the event occurred (default: current time).
    ' deltaX: number(optional) X delta in DIP for mouse wheel event (default: 0).
    ' deltaY: number(optional) Y delta in DIP for mouse wheel event (default: 0).
    ' modifiers: integer(optional) Bit field representing pressed modifier keys. Alt=1, Ctrl=2, Meta/Command=4, Shift=8
    '   (default: 0).
    ' clickCount: integer(optional) Number of times the mouse button was clicked (default: 0).

    Dim params As New Dictionary
    params("type") = CStr(pType)
    params("x") = CLng(x)
    params("y") = CLng(y)
    params("button") = CStr(button)
    If Not IsMissing(timestamp) Then params("timestamp") = CDbl(timestamp)
    If Not IsMissing(deltaX) Then params("deltaX") = CDbl(deltaX)
    If Not IsMissing(deltaY) Then params("deltaY") = CDbl(deltaY)
    If Not IsMissing(modifiers) Then params("modifiers") = CLng(modifiers)
    If Not IsMissing(clickCount) Then params("clickCount") = CLng(clickCount)

    cdp.InvokeMethod "Input.emulateTouchFromMouseEvent", params
End Sub

' Ignores input events (useful while auditing page).
Public Sub setIgnoreInputEvents( _
    ByVal ignore AS boolean _
)
    ' ignore: boolean Ignores input events processing when set to true.

    Dim params As New Dictionary
    params("ignore") = CBool(ignore)

    cdp.InvokeMethod "Input.setIgnoreInputEvents", params
End Sub

' Prevents default drag and drop behavior and instead emits `Input.dragIntercepted` events.
' Drag and drop behavior can be directly controlled via `Input.dispatchDragEvent`.
' Experimental
Public Sub setInterceptDrags( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean 

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Input.setInterceptDrags", params
End Sub

' Synthesizes a pinch gesture over a time period by issuing appropriate touch events.
' Experimental
Public Sub synthesizePinchGesture( _
    ByVal x AS Double, _
    ByVal y AS Double, _
    ByVal scaleFactor AS Double, _
    Optional ByVal relativeSpeed AS Variant, _
    Optional ByVal gestureSourceType AS Variant _
)
    ' x: number X coordinate of the start of the gesture in CSS pixels.
    ' y: number Y coordinate of the start of the gesture in CSS pixels.
    ' scaleFactor: number Relative scale factor after zooming (>1.0 zooms in, <1.0 zooms out).
    ' relativeSpeed: integer(optional) Relative pointer speed in pixels per second (default: 800).
    ' gestureSourceType: GestureSourceType(optional) Which type of input events to be generated (default: 'default', which queries the platform
    '   for the preferred input type).

    Dim params As New Dictionary
    params("x") = CDbl(x)
    params("y") = CDbl(y)
    params("scaleFactor") = CDbl(scaleFactor)
    If Not IsMissing(relativeSpeed) Then params("relativeSpeed") = CLng(relativeSpeed)
    If Not IsMissing(gestureSourceType) Then params("gestureSourceType") = CStr(gestureSourceType)

    cdp.InvokeMethod "Input.synthesizePinchGesture", params
End Sub

' Synthesizes a scroll gesture over a time period by issuing appropriate touch events.
' Experimental
Public Sub synthesizeScrollGesture( _
    ByVal x AS Double, _
    ByVal y AS Double, _
    Optional ByVal xDistance AS Variant, _
    Optional ByVal yDistance AS Variant, _
    Optional ByVal xOverscroll AS Variant, _
    Optional ByVal yOverscroll AS Variant, _
    Optional ByVal preventFling AS Variant, _
    Optional ByVal speed AS Variant, _
    Optional ByVal gestureSourceType AS Variant, _
    Optional ByVal repeatCount AS Variant, _
    Optional ByVal repeatDelayMs AS Variant, _
    Optional ByVal interactionMarkerName AS Variant _
)
    ' x: number X coordinate of the start of the gesture in CSS pixels.
    ' y: number Y coordinate of the start of the gesture in CSS pixels.
    ' xDistance: number(optional) The distance to scroll along the X axis (positive to scroll left).
    ' yDistance: number(optional) The distance to scroll along the Y axis (positive to scroll up).
    ' xOverscroll: number(optional) The number of additional pixels to scroll back along the X axis, in addition to the given
    '   distance.
    ' yOverscroll: number(optional) The number of additional pixels to scroll back along the Y axis, in addition to the given
    '   distance.
    ' preventFling: boolean(optional) Prevent fling (default: true).
    ' speed: integer(optional) Swipe speed in pixels per second (default: 800).
    ' gestureSourceType: GestureSourceType(optional) Which type of input events to be generated (default: 'default', which queries the platform
    '   for the preferred input type).
    ' repeatCount: integer(optional) The number of times to repeat the gesture (default: 0).
    ' repeatDelayMs: integer(optional) The number of milliseconds delay between each repeat. (default: 250).
    ' interactionMarkerName: string(optional) The name of the interaction markers to generate, if not empty (default: "").

    Dim params As New Dictionary
    params("x") = CDbl(x)
    params("y") = CDbl(y)
    If Not IsMissing(xDistance) Then params("xDistance") = CDbl(xDistance)
    If Not IsMissing(yDistance) Then params("yDistance") = CDbl(yDistance)
    If Not IsMissing(xOverscroll) Then params("xOverscroll") = CDbl(xOverscroll)
    If Not IsMissing(yOverscroll) Then params("yOverscroll") = CDbl(yOverscroll)
    If Not IsMissing(preventFling) Then params("preventFling") = CBool(preventFling)
    If Not IsMissing(speed) Then params("speed") = CLng(speed)
    If Not IsMissing(gestureSourceType) Then params("gestureSourceType") = CStr(gestureSourceType)
    If Not IsMissing(repeatCount) Then params("repeatCount") = CLng(repeatCount)
    If Not IsMissing(repeatDelayMs) Then params("repeatDelayMs") = CLng(repeatDelayMs)
    If Not IsMissing(interactionMarkerName) Then params("interactionMarkerName") = CStr(interactionMarkerName)

    cdp.InvokeMethod "Input.synthesizeScrollGesture", params
End Sub

' Synthesizes a tap gesture over a time period by issuing appropriate touch events.
' Experimental
Public Sub synthesizeTapGesture( _
    ByVal x AS Double, _
    ByVal y AS Double, _
    Optional ByVal duration AS Variant, _
    Optional ByVal tapCount AS Variant, _
    Optional ByVal gestureSourceType AS Variant _
)
    ' x: number X coordinate of the start of the gesture in CSS pixels.
    ' y: number Y coordinate of the start of the gesture in CSS pixels.
    ' duration: integer(optional) Duration between touchdown and touchup events in ms (default: 50).
    ' tapCount: integer(optional) Number of times to perform the tap (e.g. 2 for double tap, default: 1).
    ' gestureSourceType: GestureSourceType(optional) Which type of input events to be generated (default: 'default', which queries the platform
    '   for the preferred input type).

    Dim params As New Dictionary
    params("x") = CDbl(x)
    params("y") = CDbl(y)
    If Not IsMissing(duration) Then params("duration") = CLng(duration)
    If Not IsMissing(tapCount) Then params("tapCount") = CLng(tapCount)
    If Not IsMissing(gestureSourceType) Then params("gestureSourceType") = CStr(gestureSourceType)

    cdp.InvokeMethod "Input.synthesizeTapGesture", params
End Sub

