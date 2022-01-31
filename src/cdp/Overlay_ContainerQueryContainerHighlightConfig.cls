VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayContainerQueryContain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ContainerQueryContainerHighlightConfig"
' Overlay.ContainerQueryContainerHighlightConfig
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' containerBorder : LineStyle
' Optional
' The style of the container border.
Public containerBorder AS cdpOverlayLineStyle

' descendantBorder : LineStyle
' Optional
' The style of the descendants' borders.
Public descendantBorder AS cdpOverlayLineStyle

Public Function init(ByVal obj as Dictionary) As cdpOverlayContainerQueryContain
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("containerBorder") Then
        Set containerBorder = New cdpOverlayLineStyle
        containerBorder.init obj.Item("containerBorder")
    End If
    If obj.Exists("descendantBorder") Then
        Set descendantBorder = New cdpOverlayLineStyle
        descendantBorder.init obj.Item("descendantBorder")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not containerBorder Is Nothing Then Set dict("containerBorder") = containerBorder.asDictionary()
    If Not descendantBorder Is Nothing Then Set dict("descendantBorder") = descendantBorder.asDictionary()

    Set asDictionary = dict
End Function

