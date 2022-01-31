VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayHighlightConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "HighlightConfig"
' Overlay.HighlightConfig
' Configuration data for the highlighting of page elements.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' showInfo : boolean
' Optional
' Whether the node info tooltip should be shown (default: false).
Public showInfo AS Variant ' boolean

' showStyles : boolean
' Optional
' Whether the node styles in the tooltip (default: false).
Public showStyles AS Variant ' boolean

' showRulers : boolean
' Optional
' Whether the rulers should be shown (default: false).
Public showRulers AS Variant ' boolean

' showAccessibilityInfo : boolean
' Optional
' Whether the a11y info should be shown (default: true).
Public showAccessibilityInfo AS Variant ' boolean

' showExtensionLines : boolean
' Optional
' Whether the extension lines from node to the rulers should be shown (default: false).
Public showExtensionLines AS Variant ' boolean

' contentColor : DOM.RGBA
' Optional
' The content box highlight fill color (default: transparent).
Public contentColor AS cdpDOMRGBA

' paddingColor : DOM.RGBA
' Optional
' The padding highlight fill color (default: transparent).
Public paddingColor AS cdpDOMRGBA

' borderColor : DOM.RGBA
' Optional
' The border highlight fill color (default: transparent).
Public borderColor AS cdpDOMRGBA

' marginColor : DOM.RGBA
' Optional
' The margin highlight fill color (default: transparent).
Public marginColor AS cdpDOMRGBA

' eventTargetColor : DOM.RGBA
' Optional
' The event target element highlight fill color (default: transparent).
Public eventTargetColor AS cdpDOMRGBA

' shapeColor : DOM.RGBA
' Optional
' The shape outside fill color (default: transparent).
Public shapeColor AS cdpDOMRGBA

' shapeMarginColor : DOM.RGBA
' Optional
' The shape margin fill color (default: transparent).
Public shapeMarginColor AS cdpDOMRGBA

' cssGridColor : DOM.RGBA
' Optional
' The grid layout color (default: transparent).
Public cssGridColor AS cdpDOMRGBA

' colorFormat : ColorFormat
' Optional
' The color format used to format color styles (default: hex).
Public colorFormat AS string

' gridHighlightConfig : GridHighlightConfig
' Optional
' The grid layout highlight configuration (default: all transparent).
Public gridHighlightConfig AS cdpOverlayGridHighlightConfig

' flexContainerHighlightConfig : FlexContainerHighlightConfig
' Optional
' The flex container highlight configuration (default: all transparent).
Public flexContainerHighlightConfig AS cdpOverlayFlexContainerHighligh

' flexItemHighlightConfig : FlexItemHighlightConfig
' Optional
' The flex item highlight configuration (default: all transparent).
Public flexItemHighlightConfig AS cdpOverlayFlexItemHighlightConf

' contrastAlgorithm : ContrastAlgorithm
' Optional
' The contrast algorithm to use for the contrast ratio (default: aa).
Public contrastAlgorithm AS string

' containerQueryContainerHighlightConfig : ContainerQueryContainerHighlightConfig
' Optional
' The container query container highlight configuration (default: all transparent).
Public containerQueryContainerHighlightConfig AS cdpOverlayContainerQueryContain

Public Function init(ByVal obj as Dictionary) As cdpOverlayHighlightConfig
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("showInfo") Then Let showInfo = CBool(obj.Item("showInfo"))
    If obj.Exists("showStyles") Then Let showStyles = CBool(obj.Item("showStyles"))
    If obj.Exists("showRulers") Then Let showRulers = CBool(obj.Item("showRulers"))
    If obj.Exists("showAccessibilityInfo") Then Let showAccessibilityInfo = CBool(obj.Item("showAccessibilityInfo"))
    If obj.Exists("showExtensionLines") Then Let showExtensionLines = CBool(obj.Item("showExtensionLines"))
    If obj.Exists("contentColor") Then
        Set contentColor = New cdpDOMRGBA
        contentColor.init obj.Item("contentColor")
    End If
    If obj.Exists("paddingColor") Then
        Set paddingColor = New cdpDOMRGBA
        paddingColor.init obj.Item("paddingColor")
    End If
    If obj.Exists("borderColor") Then
        Set borderColor = New cdpDOMRGBA
        borderColor.init obj.Item("borderColor")
    End If
    If obj.Exists("marginColor") Then
        Set marginColor = New cdpDOMRGBA
        marginColor.init obj.Item("marginColor")
    End If
    If obj.Exists("eventTargetColor") Then
        Set eventTargetColor = New cdpDOMRGBA
        eventTargetColor.init obj.Item("eventTargetColor")
    End If
    If obj.Exists("shapeColor") Then
        Set shapeColor = New cdpDOMRGBA
        shapeColor.init obj.Item("shapeColor")
    End If
    If obj.Exists("shapeMarginColor") Then
        Set shapeMarginColor = New cdpDOMRGBA
        shapeMarginColor.init obj.Item("shapeMarginColor")
    End If
    If obj.Exists("cssGridColor") Then
        Set cssGridColor = New cdpDOMRGBA
        cssGridColor.init obj.Item("cssGridColor")
    End If
    If obj.Exists("colorFormat") Then Let colorFormat = CStr(obj.Item("colorFormat"))
    If obj.Exists("gridHighlightConfig") Then
        Set gridHighlightConfig = New cdpOverlayGridHighlightConfig
        gridHighlightConfig.init obj.Item("gridHighlightConfig")
    End If
    If obj.Exists("flexContainerHighlightConfig") Then
        Set flexContainerHighlightConfig = New cdpOverlayFlexContainerHighligh
        flexContainerHighlightConfig.init obj.Item("flexContainerHighlightConfig")
    End If
    If obj.Exists("flexItemHighlightConfig") Then
        Set flexItemHighlightConfig = New cdpOverlayFlexItemHighlightConf
        flexItemHighlightConfig.init obj.Item("flexItemHighlightConfig")
    End If
    If obj.Exists("contrastAlgorithm") Then Let contrastAlgorithm = CStr(obj.Item("contrastAlgorithm"))
    If obj.Exists("containerQueryContainerHighlightConfig") Then
        Set containerQueryContainerHighlightConfig = New cdpOverlayContainerQueryContain
        containerQueryContainerHighlightConfig.init obj.Item("containerQueryContainerHighlightConfig")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if Not IsEmpty(showInfo) Then dict("showInfo") = showInfo
    if Not IsEmpty(showStyles) Then dict("showStyles") = showStyles
    if Not IsEmpty(showRulers) Then dict("showRulers") = showRulers
    if Not IsEmpty(showAccessibilityInfo) Then dict("showAccessibilityInfo") = showAccessibilityInfo
    if Not IsEmpty(showExtensionLines) Then dict("showExtensionLines") = showExtensionLines
    If Not contentColor Is Nothing Then Set dict("contentColor") = contentColor.asDictionary()
    If Not paddingColor Is Nothing Then Set dict("paddingColor") = paddingColor.asDictionary()
    If Not borderColor Is Nothing Then Set dict("borderColor") = borderColor.asDictionary()
    If Not marginColor Is Nothing Then Set dict("marginColor") = marginColor.asDictionary()
    If Not eventTargetColor Is Nothing Then Set dict("eventTargetColor") = eventTargetColor.asDictionary()
    If Not shapeColor Is Nothing Then Set dict("shapeColor") = shapeColor.asDictionary()
    If Not shapeMarginColor Is Nothing Then Set dict("shapeMarginColor") = shapeMarginColor.asDictionary()
    If Not cssGridColor Is Nothing Then Set dict("cssGridColor") = cssGridColor.asDictionary()
    if colorFormat <> vbNullString Then dict("colorFormat") = colorFormat
    If Not gridHighlightConfig Is Nothing Then Set dict("gridHighlightConfig") = gridHighlightConfig.asDictionary()
    If Not flexContainerHighlightConfig Is Nothing Then Set dict("flexContainerHighlightConfig") = flexContainerHighlightConfig.asDictionary()
    If Not flexItemHighlightConfig Is Nothing Then Set dict("flexItemHighlightConfig") = flexItemHighlightConfig.asDictionary()
    if contrastAlgorithm <> vbNullString Then dict("contrastAlgorithm") = contrastAlgorithm
    If Not containerQueryContainerHighlightConfig Is Nothing Then Set dict("containerQueryContainerHighlightConfig") = containerQueryContainerHighlightConfig.asDictionary()

    Set asDictionary = dict
End Function

