VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBrowserPermissionDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PermissionDescriptor"
' Browser.PermissionDescriptor
' Definition of PermissionDescriptor defined in the Permissions API:
'   https://w3c.github.io/permissions/#dictdef-permissiondescriptor.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Name of permission.
    '   See https://cs.chromium.org/chromium/src/third_party/blink/renderer/modules/permissions/permission_descriptor.idl for valid permission names.
Public name AS string

' sysex : boolean
' Optional
' For "midi" permission, may also specify sysex control.
Public sysex AS Variant ' boolean

' userVisibleOnly : boolean
' Optional
' For "push" permission, may specify userVisibleOnly.
    '   Note that userVisibleOnly = true is the only currently supported type.
Public userVisibleOnly AS Variant ' boolean

' allowWithoutSanitization : boolean
' Optional
' For "clipboard" permission, may specify allowWithoutSanitization.
Public allowWithoutSanitization AS Variant ' boolean

' panTiltZoom : boolean
' Optional
' For "camera" permission, may specify panTiltZoom.
Public panTiltZoom AS Variant ' boolean

Public Function init(ByVal obj as Dictionary) As cdpBrowserPermissionDescriptor
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    If obj.Exists("sysex") Then Let sysex = CBool(obj.Item("sysex"))
    If obj.Exists("userVisibleOnly") Then Let userVisibleOnly = CBool(obj.Item("userVisibleOnly"))
    If obj.Exists("allowWithoutSanitization") Then Let allowWithoutSanitization = CBool(obj.Item("allowWithoutSanitization"))
    If obj.Exists("panTiltZoom") Then Let panTiltZoom = CBool(obj.Item("panTiltZoom"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    if Not IsEmpty(sysex) Then dict("sysex") = sysex
    if Not IsEmpty(userVisibleOnly) Then dict("userVisibleOnly") = userVisibleOnly
    if Not IsEmpty(allowWithoutSanitization) Then dict("allowWithoutSanitization") = allowWithoutSanitization
    if Not IsEmpty(panTiltZoom) Then dict("panTiltZoom") = panTiltZoom

    Set asDictionary = dict
End Function

