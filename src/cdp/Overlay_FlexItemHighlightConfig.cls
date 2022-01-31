VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayFlexItemHighlightConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FlexItemHighlightConfig"
' Overlay.FlexItemHighlightConfig
' Configuration data for the highlighting of Flex item elements.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' baseSizeBox : BoxStyle
' Optional
' Style of the box representing the item's base size
Public baseSizeBox AS cdpOverlayBoxStyle

' baseSizeBorder : LineStyle
' Optional
' Style of the border around the box representing the item's base size
Public baseSizeBorder AS cdpOverlayLineStyle

' flexibilityArrow : LineStyle
' Optional
' Style of the arrow representing if the item grew or shrank
Public flexibilityArrow AS cdpOverlayLineStyle

Public Function init(ByVal obj as Dictionary) As cdpOverlayFlexItemHighlightConf
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("baseSizeBox") Then
        Set baseSizeBox = New cdpOverlayBoxStyle
        baseSizeBox.init obj.Item("baseSizeBox")
    End If
    If obj.Exists("baseSizeBorder") Then
        Set baseSizeBorder = New cdpOverlayLineStyle
        baseSizeBorder.init obj.Item("baseSizeBorder")
    End If
    If obj.Exists("flexibilityArrow") Then
        Set flexibilityArrow = New cdpOverlayLineStyle
        flexibilityArrow.init obj.Item("flexibilityArrow")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not baseSizeBox Is Nothing Then Set dict("baseSizeBox") = baseSizeBox.asDictionary()
    If Not baseSizeBorder Is Nothing Then Set dict("baseSizeBorder") = baseSizeBorder.asDictionary()
    If Not flexibilityArrow Is Nothing Then Set dict("flexibilityArrow") = flexibilityArrow.asDictionary()

    Set asDictionary = dict
End Function

