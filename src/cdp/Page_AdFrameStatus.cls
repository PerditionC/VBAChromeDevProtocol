VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageAdFrameStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AdFrameStatus"
' Page.AdFrameStatus
' Indicates whether a frame has been identified as an ad and why.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' adFrameType : AdFrameType
' 
Public adFrameType AS string

' explanations : array of String
' Optional
' 
Public explanations AS Collection

Public Function init(ByVal obj as Dictionary) As cdpPageAdFrameStatus
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let adFrameType = CStr(obj.Item("adFrameType"))
    If obj.Exists("explanations") Then
        For each v in obj.Item("explanations")
            explanations.Add v
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("adFrameType") = adFrameType
    Set dict("explanations") = explanations

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set explanations = New Collection
End Sub

