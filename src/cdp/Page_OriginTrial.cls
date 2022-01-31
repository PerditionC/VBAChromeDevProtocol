VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageOriginTrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "OriginTrial"
' Page.OriginTrial
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' trialName : string
' 
Public trialName AS string

' status : OriginTrialStatus
' 
Public status AS string

' tokensWithStatus : array of cdpPageOriginTrialTokenWithStat
' 
Public tokensWithStatus AS Collection

Public Function init(ByVal obj as Dictionary) As cdpPageOriginTrial
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let trialName = CStr(obj.Item("trialName"))
    Let status = CStr(obj.Item("status"))
    For each v in obj.Item("tokensWithStatus")
        Dim obj_tokensWithStatus As cdpPageOriginTrialTokenWithStat: Set obj_tokensWithStatus = New cdpPageOriginTrialTokenWithStat
        obj_tokensWithStatus.init v
        tokensWithStatus.Add obj_tokensWithStatus
        Set obj_tokensWithStatus = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("trialName") = trialName
    dict("status") = status
    Set dict("tokensWithStatus") = tokensWithStatus

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set tokensWithStatus = New Collection
End Sub

