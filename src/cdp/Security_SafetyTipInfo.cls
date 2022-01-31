VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSecuritySafetyTipInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SafetyTipInfo"
' Security.SafetyTipInfo
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' safetyTipStatus : SafetyTipStatus
' Describes whether the page triggers any safety tips or reputation warnings. Default is unknown.
Public safetyTipStatus AS string

' safeUrl : string
' Optional
' The URL the safety tip suggested ("Did you mean?"). Only filled in for lookalike matches.
Public safeUrl AS string

Public Function init(ByVal obj as Dictionary) As cdpSecuritySafetyTipInfo
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let safetyTipStatus = CStr(obj.Item("safetyTipStatus"))
    If obj.Exists("safeUrl") Then Let safeUrl = CStr(obj.Item("safeUrl"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("safetyTipStatus") = safetyTipStatus
    if safeUrl <> vbNullString Then dict("safeUrl") = safeUrl

    Set asDictionary = dict
End Function

