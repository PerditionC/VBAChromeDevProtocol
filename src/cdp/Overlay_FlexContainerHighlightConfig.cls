VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayFlexContainerHighligh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FlexContainerHighlightConfig"
' Overlay.FlexContainerHighlightConfig
' Configuration data for the highlighting of Flex container elements.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' containerBorder : LineStyle
' Optional
' The style of the container border
Public containerBorder AS cdpOverlayLineStyle

' lineSeparator : LineStyle
' Optional
' The style of the separator between lines
Public lineSeparator AS cdpOverlayLineStyle

' itemSeparator : LineStyle
' Optional
' The style of the separator between items
Public itemSeparator AS cdpOverlayLineStyle

' mainDistributedSpace : BoxStyle
' Optional
' Style of content-distribution space on the main axis (justify-content).
Public mainDistributedSpace AS cdpOverlayBoxStyle

' crossDistributedSpace : BoxStyle
' Optional
' Style of content-distribution space on the cross axis (align-content).
Public crossDistributedSpace AS cdpOverlayBoxStyle

' rowGapSpace : BoxStyle
' Optional
' Style of empty space caused by row gaps (gap/row-gap).
Public rowGapSpace AS cdpOverlayBoxStyle

' columnGapSpace : BoxStyle
' Optional
' Style of empty space caused by columns gaps (gap/column-gap).
Public columnGapSpace AS cdpOverlayBoxStyle

' crossAlignment : LineStyle
' Optional
' Style of the self-alignment line (align-items).
Public crossAlignment AS cdpOverlayLineStyle

Public Function init(ByVal obj as Dictionary) As cdpOverlayFlexContainerHighligh
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("containerBorder") Then
        Set containerBorder = New cdpOverlayLineStyle
        containerBorder.init obj.Item("containerBorder")
    End If
    If obj.Exists("lineSeparator") Then
        Set lineSeparator = New cdpOverlayLineStyle
        lineSeparator.init obj.Item("lineSeparator")
    End If
    If obj.Exists("itemSeparator") Then
        Set itemSeparator = New cdpOverlayLineStyle
        itemSeparator.init obj.Item("itemSeparator")
    End If
    If obj.Exists("mainDistributedSpace") Then
        Set mainDistributedSpace = New cdpOverlayBoxStyle
        mainDistributedSpace.init obj.Item("mainDistributedSpace")
    End If
    If obj.Exists("crossDistributedSpace") Then
        Set crossDistributedSpace = New cdpOverlayBoxStyle
        crossDistributedSpace.init obj.Item("crossDistributedSpace")
    End If
    If obj.Exists("rowGapSpace") Then
        Set rowGapSpace = New cdpOverlayBoxStyle
        rowGapSpace.init obj.Item("rowGapSpace")
    End If
    If obj.Exists("columnGapSpace") Then
        Set columnGapSpace = New cdpOverlayBoxStyle
        columnGapSpace.init obj.Item("columnGapSpace")
    End If
    If obj.Exists("crossAlignment") Then
        Set crossAlignment = New cdpOverlayLineStyle
        crossAlignment.init obj.Item("crossAlignment")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not containerBorder Is Nothing Then Set dict("containerBorder") = containerBorder.asDictionary()
    If Not lineSeparator Is Nothing Then Set dict("lineSeparator") = lineSeparator.asDictionary()
    If Not itemSeparator Is Nothing Then Set dict("itemSeparator") = itemSeparator.asDictionary()
    If Not mainDistributedSpace Is Nothing Then Set dict("mainDistributedSpace") = mainDistributedSpace.asDictionary()
    If Not crossDistributedSpace Is Nothing Then Set dict("crossDistributedSpace") = crossDistributedSpace.asDictionary()
    If Not rowGapSpace Is Nothing Then Set dict("rowGapSpace") = rowGapSpace.asDictionary()
    If Not columnGapSpace Is Nothing Then Set dict("columnGapSpace") = columnGapSpace.asDictionary()
    If Not crossAlignment Is Nothing Then Set dict("crossAlignment") = crossAlignment.asDictionary()

    Set asDictionary = dict
End Function

