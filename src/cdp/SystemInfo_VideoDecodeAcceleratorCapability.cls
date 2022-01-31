VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSystemInfoVideoDecodeAcceler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VideoDecodeAcceleratorCapability"
' SystemInfo.VideoDecodeAcceleratorCapability
' Describes a supported video decoding profile with its associated minimum and
'   maximum resolutions.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' profile : string
' Video codec profile that is supported, e.g. VP9 Profile 2.
Public profile AS string

' maxResolution : Size
' Maximum video dimensions in pixels supported for this |profile|.
Public maxResolution AS cdpSystemInfoSize

' minResolution : Size
' Minimum video dimensions in pixels supported for this |profile|.
Public minResolution AS cdpSystemInfoSize

Public Function init(ByVal obj as Dictionary) As cdpSystemInfoVideoDecodeAcceler
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let profile = CStr(obj.Item("profile"))
    Set maxResolution = New cdpSystemInfoSize
    maxResolution.init obj.Item("maxResolution")
    Set minResolution = New cdpSystemInfoSize
    minResolution.init obj.Item("minResolution")

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("profile") = profile
    Set dict("maxResolution") = maxResolution.asDictionary()
    Set dict("minResolution") = minResolution.asDictionary()

    Set asDictionary = dict
End Function

