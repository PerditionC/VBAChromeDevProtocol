VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotRareBooleanData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "RareBooleanData"
' DOMSnapshot.RareBooleanData
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' index : array of integer
' 
Public index AS Collection

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotRareBooleanData
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("index")
        index.Add v
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("index") = index

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set index = New Collection
End Sub

