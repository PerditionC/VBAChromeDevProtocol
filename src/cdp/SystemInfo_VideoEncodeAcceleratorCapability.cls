VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSystemInfoVideoEncodeAcceler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VideoEncodeAcceleratorCapability"
' SystemInfo.VideoEncodeAcceleratorCapability
' Describes a supported video encoding profile with its associated maximum
'   resolution and maximum framerate.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' profile : string
' Video codec profile that is supported, e.g H264 Main.
Public profile AS string

' maxResolution : Size
' Maximum video dimensions in pixels supported for this |profile|.
Public maxResolution AS cdpSystemInfoSize

' maxFramerateNumerator : integer
' Maximum encoding framerate in frames per second supported for this
    '   |profile|, as fraction's numerator and denominator, e.g. 24/1 fps,
    '   24000/1001 fps, etc.
Public maxFramerateNumerator AS Long

' maxFramerateDenominator : integer
' 
Public maxFramerateDenominator AS Long

Public Function init(ByVal obj as Dictionary) As cdpSystemInfoVideoEncodeAcceler
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let profile = CStr(obj.Item("profile"))
    Set maxResolution = New cdpSystemInfoSize
    maxResolution.init obj.Item("maxResolution")
    Let maxFramerateNumerator = CLng(obj.Item("maxFramerateNumerator"))
    Let maxFramerateDenominator = CLng(obj.Item("maxFramerateDenominator"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("profile") = profile
    Set dict("maxResolution") = maxResolution.asDictionary()
    dict("maxFramerateNumerator") = maxFramerateNumerator
    dict("maxFramerateDenominator") = maxFramerateDenominator

    Set asDictionary = dict
End Function

