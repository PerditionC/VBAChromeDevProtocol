VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPagePermissionsPolicyBlockLo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PermissionsPolicyBlockLocator"
' Page.PermissionsPolicyBlockLocator
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' frameId : FrameId
' 
Public frameId AS string

' blockReason : PermissionsPolicyBlockReason
' 
Public blockReason AS string

Public Function init(ByVal obj as Dictionary) As cdpPagePermissionsPolicyBlockLo
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let frameId = CStr(obj.Item("frameId"))
    Let blockReason = CStr(obj.Item("blockReason"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("frameId") = frameId
    dict("blockReason") = blockReason

    Set asDictionary = dict
End Function

