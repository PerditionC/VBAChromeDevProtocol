VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpInputTouchPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "TouchPoint"
' Input.TouchPoint
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' x : number
' X coordinate of the event relative to the main frame's viewport in CSS pixels.
Public x AS Double

' y : number
' Y coordinate of the event relative to the main frame's viewport in CSS pixels. 0 refers to
    '   the top of the viewport and Y increases as it proceeds towards the bottom of the viewport.
Public y AS Double

' radiusX : number
' Optional
' X radius of the touch area (default: 1.0).
Public radiusX AS Double

' radiusY : number
' Optional
' Y radius of the touch area (default: 1.0).
Public radiusY AS Double

' rotationAngle : number
' Optional
' Rotation angle (default: 0.0).
Public rotationAngle AS Double

' force : number
' Optional
' Force (default: 1.0).
Public force AS Double

' tangentialPressure : number
' Optional
' The normalized tangential pressure, which has a range of [-1,1] (default: 0).
Public tangentialPressure AS Double

' tiltX : integer
' Optional
' The plane angle between the Y-Z plane and the plane containing both the stylus axis and the Y axis, in degrees of the range [-90,90], a positive tiltX is to the right (default: 0)
Public tiltX AS Long

' tiltY : integer
' Optional
' The plane angle between the X-Z plane and the plane containing both the stylus axis and the X axis, in degrees of the range [-90,90], a positive tiltY is towards the user (default: 0).
Public tiltY AS Long

' twist : integer
' Optional
' The clockwise rotation of a pen stylus around its own major axis, in degrees in the range [0,359] (default: 0).
Public twist AS Long

' id : number
' Optional
' Identifier used to track touch sources between events, must be unique within an event.
Public id AS Double

Public Function init(ByVal obj as Dictionary) As cdpInputTouchPoint
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let x = CDbl(obj.Item("x"))
    Let y = CDbl(obj.Item("y"))
    If obj.Exists("radiusX") Then Let radiusX = CDbl(obj.Item("radiusX"))
    If obj.Exists("radiusY") Then Let radiusY = CDbl(obj.Item("radiusY"))
    If obj.Exists("rotationAngle") Then Let rotationAngle = CDbl(obj.Item("rotationAngle"))
    If obj.Exists("force") Then Let force = CDbl(obj.Item("force"))
    If obj.Exists("tangentialPressure") Then Let tangentialPressure = CDbl(obj.Item("tangentialPressure"))
    If obj.Exists("tiltX") Then Let tiltX = CLng(obj.Item("tiltX"))
    If obj.Exists("tiltY") Then Let tiltY = CLng(obj.Item("tiltY"))
    If obj.Exists("twist") Then Let twist = CLng(obj.Item("twist"))
    If obj.Exists("id") Then Let id = CDbl(obj.Item("id"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("x") = x
    dict("y") = y
    dict("radiusX") = radiusX
    dict("radiusY") = radiusY
    dict("rotationAngle") = rotationAngle
    dict("force") = force
    dict("tangentialPressure") = tangentialPressure
    dict("tiltX") = tiltX
    dict("tiltY") = tiltY
    dict("twist") = twist
    dict("id") = id

    Set asDictionary = dict
End Function

