VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpEmulationDisplayFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DisplayFeature"
' Emulation.DisplayFeature
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' orientation : string
' Orientation of a display feature in relation to screen
Public orientation AS string

' offset : integer
' The offset from the screen origin in either the x (for vertical
    '   orientation) or y (for horizontal orientation) direction.
Public offset AS Long

' maskLength : integer
' A display feature may mask content such that it is not physically
    '   displayed - this length along with the offset describes this area.
    '   A display feature that only splits content will have a 0 mask_length.
Public maskLength AS Long

Public Function init(ByVal obj as Dictionary) As cdpEmulationDisplayFeature
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let orientation = CStr(obj.Item("orientation"))
    Let offset = CLng(obj.Item("offset"))
    Let maskLength = CLng(obj.Item("maskLength"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("orientation") = orientation
    dict("offset") = offset
    dict("maskLength") = maskLength

    Set asDictionary = dict
End Function

