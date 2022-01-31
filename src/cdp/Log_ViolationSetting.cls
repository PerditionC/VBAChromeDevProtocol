VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpLogViolationSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ViolationSetting"
' Log.ViolationSetting
' Violation configuration setting.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Violation type.
Public name AS string

' threshold : number
' Time threshold to trigger upon.
Public threshold AS Double

Public Function init(ByVal obj as Dictionary) As cdpLogViolationSetting
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    Let threshold = CDbl(obj.Item("threshold"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("threshold") = threshold

    Set asDictionary = dict
End Function

