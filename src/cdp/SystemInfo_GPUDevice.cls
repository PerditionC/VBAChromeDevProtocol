VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSystemInfoGPUDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "GPUDevice"
' SystemInfo.GPUDevice
' Describes a single graphics processor (GPU).
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' vendorId : number
' PCI ID of the GPU vendor, if available; 0 otherwise.
Public vendorId AS Double

' deviceId : number
' PCI ID of the GPU device, if available; 0 otherwise.
Public deviceId AS Double

' subSysId : number
' Optional
' Sub sys ID of the GPU, only available on Windows.
Public subSysId AS Double

' revision : number
' Optional
' Revision of the GPU, only available on Windows.
Public revision AS Double

' vendorString : string
' String description of the GPU vendor, if the PCI ID is not available.
Public vendorString AS string

' deviceString : string
' String description of the GPU device, if the PCI ID is not available.
Public deviceString AS string

' driverVendor : string
' String description of the GPU driver vendor.
Public driverVendor AS string

' driverVersion : string
' String description of the GPU driver version.
Public driverVersion AS string

Public Function init(ByVal obj as Dictionary) As cdpSystemInfoGPUDevice
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let vendorId = CDbl(obj.Item("vendorId"))
    Let deviceId = CDbl(obj.Item("deviceId"))
    If obj.Exists("subSysId") Then Let subSysId = CDbl(obj.Item("subSysId"))
    If obj.Exists("revision") Then Let revision = CDbl(obj.Item("revision"))
    Let vendorString = CStr(obj.Item("vendorString"))
    Let deviceString = CStr(obj.Item("deviceString"))
    Let driverVendor = CStr(obj.Item("driverVendor"))
    Let driverVersion = CStr(obj.Item("driverVersion"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("vendorId") = vendorId
    dict("deviceId") = deviceId
    dict("subSysId") = subSysId
    dict("revision") = revision
    dict("vendorString") = vendorString
    dict("deviceString") = deviceString
    dict("driverVendor") = driverVendor
    dict("driverVersion") = driverVersion

    Set asDictionary = dict
End Function

