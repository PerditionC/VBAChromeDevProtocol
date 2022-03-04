VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageNavigationEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "NavigationEntry"
' Page.NavigationEntry
' Navigation history entry.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' id : integer
' Unique id of the navigation history entry.
Public id AS Long

' url : string
' URL of the navigation history entry.
Public url AS string

' userTypedURL : string
' URL that the user typed in the url bar.
Public userTypedURL AS string

' title : string
' Title of the navigation history entry.
Public title AS string

' transitionType : TransitionType
' Transition type.
Public transitionType AS string

Public Function init(ByVal obj as Dictionary) As cdpPageNavigationEntry
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let id = CLng(obj.Item("id"))
    Let url = CStr(obj.Item("url"))
    Let userTypedURL = CStr(obj.Item("userTypedURL"))
    Let title = CStr(obj.Item("title"))
    Let transitionType = CStr(obj.Item("transitionType"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("id") = id
    dict("url") = url
    dict("userTypedURL") = userTypedURL
    dict("title") = title
    dict("transitionType") = transitionType

    Set asDictionary = dict
End Function

