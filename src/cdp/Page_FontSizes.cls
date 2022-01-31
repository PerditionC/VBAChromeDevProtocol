VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageFontSizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Page.FontSizes
' Default font sizes.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' standard : integer
' Optional
' Default standard font size.
Public standard AS integer

' fixed : integer
' Optional
' Default fixed font size.
Public fixed AS integer

Public Function init(ByVal obj as Dictionary) As cdpPageFontSizes

    If obj.Exists("standard") Then Let standard = CInt(obj.Item("standard"))
    If obj.Exists("fixed") Then Let fixed = CInt(obj.Item("fixed"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("standard") = standard
    dict("fixed") = fixed

    Set asDictionary = dict
End Function
