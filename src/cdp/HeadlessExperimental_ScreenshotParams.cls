VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpHeadlessExperimentalScreensh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScreenshotParams"
' HeadlessExperimental.ScreenshotParams
' Encoding options for a screenshot.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' format : string
' Optional
' Image compression format (defaults to png).
Public format AS string

' quality : integer
' Optional
' Compression quality from range [0..100] (jpeg only).
Public quality AS Long

Public Function init(ByVal obj as Dictionary) As cdpHeadlessExperimentalScreensh
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("format") Then Let format = CStr(obj.Item("format"))
    If obj.Exists("quality") Then Let quality = CLng(obj.Item("quality"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if format <> vbNullString Then dict("format") = format
    dict("quality") = quality

    Set asDictionary = dict
End Function

