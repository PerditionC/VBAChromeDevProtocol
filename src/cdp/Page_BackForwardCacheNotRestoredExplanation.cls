VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageBackForwardCacheNotResto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "BackForwardCacheNotRestoredExplanation"
' Page.BackForwardCacheNotRestoredExplanation
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : BackForwardCacheNotRestoredReasonType
' Type of the reason
Public fieldType AS string

' reason : BackForwardCacheNotRestoredReason
' Not restored reason
Public reason AS string

Public Function init(ByVal obj as Dictionary) As cdpPageBackForwardCacheNotResto
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let fieldType = CStr(obj.Item("type"))
    Let reason = CStr(obj.Item("reason"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    dict("reason") = reason

    Set asDictionary = dict
End Function

