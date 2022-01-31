VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBrowserBounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Browser.Bounds
' Browser window bounds information
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' left : integer
' Optional
' The offset from the left edge of the screen to the window in pixels.
Public left AS integer

' top : integer
' Optional
' The offset from the top edge of the screen to the window in pixels.
Public top AS integer

' width : integer
' Optional
' The window width in pixels.
Public width AS integer

' height : integer
' Optional
' The window height in pixels.
Public height AS integer

' windowState : WindowState
' Optional
' The window state. Default to normal.
Public windowState AS string

Public Function init(ByVal obj as Dictionary) As cdpBrowserBounds

    If obj.Exists("left") Then Let left = CInt(obj.Item("left"))
    If obj.Exists("top") Then Let top = CInt(obj.Item("top"))
    If obj.Exists("width") Then Let width = CInt(obj.Item("width"))
    If obj.Exists("height") Then Let height = CInt(obj.Item("height"))
    If obj.Exists("windowState") Then Let windowState = CStr(obj.Item("windowState"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("left") = left
    dict("top") = top
    dict("width") = width
    dict("height") = height
    if windowState <> vbNullString Then dict("windowState") = windowState

    Set asDictionary = dict
End Function

