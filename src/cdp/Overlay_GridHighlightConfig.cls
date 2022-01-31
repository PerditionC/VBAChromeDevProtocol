VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpOverlayGridHighlightConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "GridHighlightConfig"
' Overlay.GridHighlightConfig
' Configuration data for the highlighting of Grid elements.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' showGridExtensionLines : boolean
' Optional
' Whether the extension lines from grid cells to the rulers should be shown (default: false).
Public showGridExtensionLines AS Variant ' boolean

' showPositiveLineNumbers : boolean
' Optional
' Show Positive line number labels (default: false).
Public showPositiveLineNumbers AS Variant ' boolean

' showNegativeLineNumbers : boolean
' Optional
' Show Negative line number labels (default: false).
Public showNegativeLineNumbers AS Variant ' boolean

' showAreaNames : boolean
' Optional
' Show area name labels (default: false).
Public showAreaNames AS Variant ' boolean

' showLineNames : boolean
' Optional
' Show line name labels (default: false).
Public showLineNames AS Variant ' boolean

' showTrackSizes : boolean
' Optional
' Show track size labels (default: false).
Public showTrackSizes AS Variant ' boolean

' gridBorderColor : DOM.RGBA
' Optional
' The grid container border highlight color (default: transparent).
Public gridBorderColor AS cdpDOMRGBA

' cellBorderColor : DOM.RGBA
' Optional
' The cell border color (default: transparent). Deprecated, please use rowLineColor and columnLineColor instead.
Public cellBorderColor AS cdpDOMRGBA

' rowLineColor : DOM.RGBA
' Optional
' The row line color (default: transparent).
Public rowLineColor AS cdpDOMRGBA

' columnLineColor : DOM.RGBA
' Optional
' The column line color (default: transparent).
Public columnLineColor AS cdpDOMRGBA

' gridBorderDash : boolean
' Optional
' Whether the grid border is dashed (default: false).
Public gridBorderDash AS Variant ' boolean

' cellBorderDash : boolean
' Optional
' Whether the cell border is dashed (default: false). Deprecated, please us rowLineDash and columnLineDash instead.
Public cellBorderDash AS Variant ' boolean

' rowLineDash : boolean
' Optional
' Whether row lines are dashed (default: false).
Public rowLineDash AS Variant ' boolean

' columnLineDash : boolean
' Optional
' Whether column lines are dashed (default: false).
Public columnLineDash AS Variant ' boolean

' rowGapColor : DOM.RGBA
' Optional
' The row gap highlight fill color (default: transparent).
Public rowGapColor AS cdpDOMRGBA

' rowHatchColor : DOM.RGBA
' Optional
' The row gap hatching fill color (default: transparent).
Public rowHatchColor AS cdpDOMRGBA

' columnGapColor : DOM.RGBA
' Optional
' The column gap highlight fill color (default: transparent).
Public columnGapColor AS cdpDOMRGBA

' columnHatchColor : DOM.RGBA
' Optional
' The column gap hatching fill color (default: transparent).
Public columnHatchColor AS cdpDOMRGBA

' areaBorderColor : DOM.RGBA
' Optional
' The named grid areas border color (Default: transparent).
Public areaBorderColor AS cdpDOMRGBA

' gridBackgroundColor : DOM.RGBA
' Optional
' The grid container background color (Default: transparent).
Public gridBackgroundColor AS cdpDOMRGBA

Public Function init(ByVal obj as Dictionary) As cdpOverlayGridHighlightConfig
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("showGridExtensionLines") Then Let showGridExtensionLines = CBool(obj.Item("showGridExtensionLines"))
    If obj.Exists("showPositiveLineNumbers") Then Let showPositiveLineNumbers = CBool(obj.Item("showPositiveLineNumbers"))
    If obj.Exists("showNegativeLineNumbers") Then Let showNegativeLineNumbers = CBool(obj.Item("showNegativeLineNumbers"))
    If obj.Exists("showAreaNames") Then Let showAreaNames = CBool(obj.Item("showAreaNames"))
    If obj.Exists("showLineNames") Then Let showLineNames = CBool(obj.Item("showLineNames"))
    If obj.Exists("showTrackSizes") Then Let showTrackSizes = CBool(obj.Item("showTrackSizes"))
    If obj.Exists("gridBorderColor") Then
        Set gridBorderColor = New cdpDOMRGBA
        gridBorderColor.init obj.Item("gridBorderColor")
    End If
    If obj.Exists("cellBorderColor") Then
        Set cellBorderColor = New cdpDOMRGBA
        cellBorderColor.init obj.Item("cellBorderColor")
    End If
    If obj.Exists("rowLineColor") Then
        Set rowLineColor = New cdpDOMRGBA
        rowLineColor.init obj.Item("rowLineColor")
    End If
    If obj.Exists("columnLineColor") Then
        Set columnLineColor = New cdpDOMRGBA
        columnLineColor.init obj.Item("columnLineColor")
    End If
    If obj.Exists("gridBorderDash") Then Let gridBorderDash = CBool(obj.Item("gridBorderDash"))
    If obj.Exists("cellBorderDash") Then Let cellBorderDash = CBool(obj.Item("cellBorderDash"))
    If obj.Exists("rowLineDash") Then Let rowLineDash = CBool(obj.Item("rowLineDash"))
    If obj.Exists("columnLineDash") Then Let columnLineDash = CBool(obj.Item("columnLineDash"))
    If obj.Exists("rowGapColor") Then
        Set rowGapColor = New cdpDOMRGBA
        rowGapColor.init obj.Item("rowGapColor")
    End If
    If obj.Exists("rowHatchColor") Then
        Set rowHatchColor = New cdpDOMRGBA
        rowHatchColor.init obj.Item("rowHatchColor")
    End If
    If obj.Exists("columnGapColor") Then
        Set columnGapColor = New cdpDOMRGBA
        columnGapColor.init obj.Item("columnGapColor")
    End If
    If obj.Exists("columnHatchColor") Then
        Set columnHatchColor = New cdpDOMRGBA
        columnHatchColor.init obj.Item("columnHatchColor")
    End If
    If obj.Exists("areaBorderColor") Then
        Set areaBorderColor = New cdpDOMRGBA
        areaBorderColor.init obj.Item("areaBorderColor")
    End If
    If obj.Exists("gridBackgroundColor") Then
        Set gridBackgroundColor = New cdpDOMRGBA
        gridBackgroundColor.init obj.Item("gridBackgroundColor")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if Not IsEmpty(showGridExtensionLines) Then dict("showGridExtensionLines") = showGridExtensionLines
    if Not IsEmpty(showPositiveLineNumbers) Then dict("showPositiveLineNumbers") = showPositiveLineNumbers
    if Not IsEmpty(showNegativeLineNumbers) Then dict("showNegativeLineNumbers") = showNegativeLineNumbers
    if Not IsEmpty(showAreaNames) Then dict("showAreaNames") = showAreaNames
    if Not IsEmpty(showLineNames) Then dict("showLineNames") = showLineNames
    if Not IsEmpty(showTrackSizes) Then dict("showTrackSizes") = showTrackSizes
    If Not gridBorderColor Is Nothing Then Set dict("gridBorderColor") = gridBorderColor.asDictionary()
    If Not cellBorderColor Is Nothing Then Set dict("cellBorderColor") = cellBorderColor.asDictionary()
    If Not rowLineColor Is Nothing Then Set dict("rowLineColor") = rowLineColor.asDictionary()
    If Not columnLineColor Is Nothing Then Set dict("columnLineColor") = columnLineColor.asDictionary()
    if Not IsEmpty(gridBorderDash) Then dict("gridBorderDash") = gridBorderDash
    if Not IsEmpty(cellBorderDash) Then dict("cellBorderDash") = cellBorderDash
    if Not IsEmpty(rowLineDash) Then dict("rowLineDash") = rowLineDash
    if Not IsEmpty(columnLineDash) Then dict("columnLineDash") = columnLineDash
    If Not rowGapColor Is Nothing Then Set dict("rowGapColor") = rowGapColor.asDictionary()
    If Not rowHatchColor Is Nothing Then Set dict("rowHatchColor") = rowHatchColor.asDictionary()
    If Not columnGapColor Is Nothing Then Set dict("columnGapColor") = columnGapColor.asDictionary()
    If Not columnHatchColor Is Nothing Then Set dict("columnHatchColor") = columnHatchColor.asDictionary()
    If Not areaBorderColor Is Nothing Then Set dict("areaBorderColor") = areaBorderColor.asDictionary()
    If Not gridBackgroundColor Is Nothing Then Set dict("gridBackgroundColor") = gridBackgroundColor.asDictionary()

    Set asDictionary = dict
End Function

