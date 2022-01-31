VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSystemInfoGPUInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "GPUInfo"
' SystemInfo.GPUInfo
' Provides information about the GPU(s) on the system.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' devices : array of cdpSystemInfoGPUDevice
' The graphics devices on the system. Element 0 is the primary GPU.
Public devices AS Collection

' auxAttributes : object
' Optional
' An optional dictionary of additional GPU related attributes.
Public auxAttributes AS Dictionary

' featureStatus : object
' Optional
' An optional dictionary of graphics features and their status.
Public featureStatus AS Dictionary

' driverBugWorkarounds : array of string
' An optional array of GPU driver bug workarounds.
Public driverBugWorkarounds AS Collection

' videoDecoding : array of cdpSystemInfoVideoDecodeAcceler
' Supported accelerated video decoding capabilities.
Public videoDecoding AS Collection

' videoEncoding : array of cdpSystemInfoVideoEncodeAcceler
' Supported accelerated video encoding capabilities.
Public videoEncoding AS Collection

' imageDecoding : array of cdpSystemInfoImageDecodeAcceler
' Supported accelerated image decoding capabilities.
Public imageDecoding AS Collection

Public Function init(ByVal obj as Dictionary) As cdpSystemInfoGPUInfo
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("devices")
        Dim obj_devices As cdpSystemInfoGPUDevice: Set obj_devices = New cdpSystemInfoGPUDevice
        obj_devices.init v
        devices.Add obj_devices
        Set obj_devices = Nothing
    Next v
    If obj.Exists("auxAttributes") Then
        Set auxAttributes = obj.Item("auxAttributes")
    End If
    If obj.Exists("featureStatus") Then
        Set featureStatus = obj.Item("featureStatus")
    End If
    For each v in obj.Item("driverBugWorkarounds")
        driverBugWorkarounds.Add v
    Next v
    For each v in obj.Item("videoDecoding")
        Dim obj_videoDecoding As cdpSystemInfoVideoDecodeAcceler: Set obj_videoDecoding = New cdpSystemInfoVideoDecodeAcceler
        obj_videoDecoding.init v
        videoDecoding.Add obj_videoDecoding
        Set obj_videoDecoding = Nothing
    Next v
    For each v in obj.Item("videoEncoding")
        Dim obj_videoEncoding As cdpSystemInfoVideoEncodeAcceler: Set obj_videoEncoding = New cdpSystemInfoVideoEncodeAcceler
        obj_videoEncoding.init v
        videoEncoding.Add obj_videoEncoding
        Set obj_videoEncoding = Nothing
    Next v
    For each v in obj.Item("imageDecoding")
        Dim obj_imageDecoding As cdpSystemInfoImageDecodeAcceler: Set obj_imageDecoding = New cdpSystemInfoImageDecodeAcceler
        obj_imageDecoding.init v
        imageDecoding.Add obj_imageDecoding
        Set obj_imageDecoding = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("devices") = devices
    If Not auxAttributes Is Nothing Then Set dict("auxAttributes") = auxAttributes
    If Not featureStatus Is Nothing Then Set dict("featureStatus") = featureStatus
    Set dict("driverBugWorkarounds") = driverBugWorkarounds
    Set dict("videoDecoding") = videoDecoding
    Set dict("videoEncoding") = videoEncoding
    Set dict("imageDecoding") = imageDecoding

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set devices = New Collection
    Set driverBugWorkarounds = New Collection
    Set videoDecoding = New Collection
    Set videoEncoding = New Collection
    Set imageDecoding = New Collection
End Sub

