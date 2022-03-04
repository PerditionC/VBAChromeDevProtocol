VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSystemInfoSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Size"
' SystemInfo.Size
' Describes the width and height dimensions of an entity.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' width : integer
' Width in pixels.
Public width AS Long

' height : integer
' Height in pixels.
Public height AS Long

Public Function init(ByVal obj as Dictionary) As cdpSystemInfoSize
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let width = CLng(obj.Item("width"))
    Let height = CLng(obj.Item("height"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("width") = width
    dict("height") = height

    Set asDictionary = dict
End Function

