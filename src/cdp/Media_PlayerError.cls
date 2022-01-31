VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpMediaPlayerError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PlayerError"
' Media.PlayerError
' Corresponds to kMediaError
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : string
' 
Public fieldType AS string

' errorCode : string
' When this switches to using media::Status instead of PipelineStatus
    '   we can remove "errorCode" and replace it with the fields from
    '   a Status instance. This also seems like a duplicate of the error
    '   level enum - there is a todo bug to have that level removed and
    '   use this instead. (crbug.com/1068454)
Public errorCode AS string

Public Function init(ByVal obj as Dictionary) As cdpMediaPlayerError
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let fieldType = CStr(obj.Item("type"))
    Let errorCode = CStr(obj.Item("errorCode"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    dict("errorCode") = errorCode

    Set asDictionary = dict
End Function

