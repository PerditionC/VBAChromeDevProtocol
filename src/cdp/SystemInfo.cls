VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpSystemInfo"
' SystemInfo [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Describes a single graphics processor (GPU).
' GPUDevice AS object

' Describes the width and height dimensions of an entity.
' Size AS object

' Describes a supported video decoding profile with its associated minimum and
'   maximum resolutions.
' VideoDecodeAcceleratorCapability AS object

' Describes a supported video encoding profile with its associated maximum
'   resolution and maximum framerate.
' VideoEncodeAcceleratorCapability AS object

' YUV subsampling type of the pixels of a given image.
' SubsamplingFormat AS string
Public Enum SubsamplingFormat
    SF_yuv420
    SF_yuv422
    SF_yuv444
End Enum

' Image format of a given image.
' ImageType AS string
Public Enum ImageType
    IT_jpeg
    IT_webp
    IT_unknown
End Enum

' Describes a supported image decoding profile with its associated minimum and
'   maximum resolutions and subsampling.
' ImageDecodeAcceleratorCapability AS object

' Provides information about the GPU(s) on the system.
' GPUInfo AS object

' Represents process info.
' ProcessInfo AS object


Public Function SubsamplingFormatToString(ByVal sf As SubsamplingFormat) As String
    Dim retVal As String
    Select Case sf
        Case SF_yuv420
            retVal = "yuv420"
        Case SF_yuv422
            retVal = "yuv422"
        Case SF_yuv444
            retVal = "yuv444"
        Case Else
            Debug.Print "Warning, unknown value " & sf
    End Select
    SubsamplingFormatToString = retVal
End Function

Public Function StringToSubsamplingFormat(ByVal s As String) As SubsamplingFormat
    Dim retVal As SubsamplingFormat
    Select Case s
        Case "yuv420"
            retVal = SF_yuv420
        Case "yuv422"
            retVal = SF_yuv422
        Case "yuv444"
            retVal = SF_yuv444
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToSubsamplingFormat = retVal
End Function


Public Function ImageTypeToString(ByVal it As ImageType) As String
    Dim retVal As String
    Select Case it
        Case IT_jpeg
            retVal = "jpeg"
        Case IT_webp
            retVal = "webp"
        Case IT_unknown
            retVal = "unknown"
        Case Else
            Debug.Print "Warning, unknown value " & it
    End Select
    ImageTypeToString = retVal
End Function

Public Function StringToImageType(ByVal s As String) As ImageType
    Dim retVal As ImageType
    Select Case s
        Case "jpeg"
            retVal = IT_jpeg
        Case "webp"
            retVal = IT_webp
        Case "unknown"
            retVal = IT_unknown
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToImageType = retVal
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

' Returns information about the system.
' Returns:  - 
Public Function getInfo() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("SystemInfo.getInfo", params)
    If cdp.ErrorCode = 0 Then
        Set getInfo = results
    End If
End Function

' Returns information about all running processes.
' Returns: processInfo - An array of process info blocks.
Public Function getProcessInfo() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("SystemInfo.getProcessInfo", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("processInfo") Then Set getProcessInfo = results("processInfo")
    End If
End Function

