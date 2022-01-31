VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeEntryPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "EntryPreview"
' Runtime.EntryPreview
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' key : ObjectPreview
' Optional
' Preview of the key. Specified for map-like collection entries.
Public key AS cdpRuntimeObjectPreview

' value : ObjectPreview
' Preview of the value.
Public value AS cdpRuntimeObjectPreview

Public Function init(ByVal obj as Dictionary) As cdpRuntimeEntryPreview
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("key") Then
        Set key = New cdpRuntimeObjectPreview
        key.init obj.Item("key")
    End If
    Set value = New cdpRuntimeObjectPreview
    value.init obj.Item("value")

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not key Is Nothing Then Set dict("key") = key.asDictionary()
    Set dict("value") = value.asDictionary()

    Set asDictionary = dict
End Function

