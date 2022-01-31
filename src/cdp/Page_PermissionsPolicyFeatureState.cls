VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPagePermissionsPolicyFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PermissionsPolicyFeatureState"
' Page.PermissionsPolicyFeatureState
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' feature : PermissionsPolicyFeature
' 
Public feature AS string

' allowed : boolean
' 
Public allowed AS boolean

' locator : PermissionsPolicyBlockLocator
' Optional
' 
Public locator AS cdpPagePermissionsPolicyBlockLo

Public Function init(ByVal obj as Dictionary) As cdpPagePermissionsPolicyFeature
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let feature = CStr(obj.Item("feature"))
    Let allowed = CBool(obj.Item("allowed"))
    If obj.Exists("locator") Then
        Set locator = New cdpPagePermissionsPolicyBlockLo
        locator.init obj.Item("locator")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("feature") = feature
    dict("allowed") = allowed
    If Not locator Is Nothing Then Set dict("locator") = locator.asDictionary()

    Set asDictionary = dict
End Function

