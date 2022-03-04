VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpCSS"
' CSS [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' 
' StyleSheetId AS string

' Stylesheet type: "injected" for stylesheets injected via extension, "user-agent" for user-agent
'   stylesheets, "inspector" for stylesheets created by the inspector (i.e. those holding the "via
'   inspector" rules), "regular" for regular stylesheets.
' StyleSheetOrigin AS string
Public Enum StyleSheetOrigin
    SSO_injected
    SSO_useragent
    SSO_inspector
    SSO_regular
End Enum

' CSS rule collection for a single pseudo style.
' PseudoElementMatches AS object

' Inherited CSS rule collection from ancestor node.
' InheritedStyleEntry AS object

' Match data for a CSS rule.
' RuleMatch AS object

' Data for a simple selector (these are delimited by commas in a selector list).
' Value AS object

' Selector list data.
' SelectorList AS object

' CSS stylesheet metainformation.
' CSSStyleSheetHeader AS object

' CSS rule representation.
' CSSRule AS object

' CSS coverage information.
' RuleUsage AS object

' Text range within a resource. All numbers are zero-based.
' SourceRange AS object

' 
' ShorthandEntry AS object

' 
' CSSComputedStyleProperty AS object

' CSS style representation.
' CSSStyle AS object

' CSS property declaration data.
' CSSProperty AS object

' CSS media rule descriptor.
' CSSMedia AS object

' Media query descriptor.
' MediaQuery AS object

' Media query expression descriptor.
' MediaQueryExpression AS object

' CSS container query rule descriptor.
' CSSContainerQuery AS object

' Information about amount of glyphs that were rendered with given font.
' PlatformFontUsage AS object

' Information about font variation axes for variable fonts
' FontVariationAxis AS object

' Properties of a web font: https://www.w3.org/TR/2008/REC-CSS2-20080411/fonts.html#font-descriptions
'   and additional information such as platformFontFamily and fontVariationAxes.
' FontFace AS object

' CSS keyframes rule representation.
' CSSKeyframesRule AS object

' CSS keyframe rule representation.
' CSSKeyframeRule AS object

' A descriptor of operation to mutate style declaration text.
' StyleDeclarationEdit AS object


Public Function StyleSheetOriginToString(ByVal sso As StyleSheetOrigin) As String
    Dim retVal As String
    Select Case sso
        Case SSO_injected
            retVal = "injected"
        Case SSO_useragent
            retVal = "user-agent"
        Case SSO_inspector
            retVal = "inspector"
        Case SSO_regular
            retVal = "regular"
        Case Else
            Debug.Print "Warning, unknown value " & sso
    End Select
    StyleSheetOriginToString = retVal
End Function

Public Function StringToStyleSheetOrigin(ByVal s As String) As StyleSheetOrigin
    Dim retVal As StyleSheetOrigin
    Select Case s
        Case "injected"
            retVal = SSO_injected
        Case "user-agent"
            retVal = SSO_useragent
        Case "inspector"
            retVal = SSO_inspector
        Case "regular"
            retVal = SSO_regular
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToStyleSheetOrigin = retVal
End Function


Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' Inserts a new rule with the given `ruleText` in a stylesheet with given `styleSheetId`, at the
' position specified by `location`.
' Returns: rule - The newly created rule.
Public Function addRule( _
    ByVal styleSheetId AS string, _
    ByVal ruleText AS string, _
    ByVal location AS Dictionary _
) AS Dictionary
    ' styleSheetId: StyleSheetId The css style sheet identifier where a new rule should be inserted.
    ' ruleText: string The text of a new rule.
    ' location: SourceRange Text position of a new rule in the target style sheet.

    Dim params As New Dictionary
    params("styleSheetId") = CStr(styleSheetId)
    params("ruleText") = CStr(ruleText)
    Set params("location") = location

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.addRule", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("rule") Then Set addRule = results("rule")
    End If
End Function

' Returns all class names from specified stylesheet.
' Returns: classNames - Class name list.
Public Function collectClassNames( _
    ByVal styleSheetId AS string _
) AS Collection
    ' styleSheetId: StyleSheetId 

    Dim params As New Dictionary
    params("styleSheetId") = CStr(styleSheetId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.collectClassNames", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("classNames") Then Set collectClassNames = results("classNames")
    End If
End Function

' Creates a new special "via-inspector" stylesheet in the frame with given `frameId`.
' Returns: styleSheetId - Identifier of the created "via-inspector" stylesheet.
Public Function createStyleSheet( _
    ByVal frameId AS string _
) AS string
    ' frameId: Page.FrameId Identifier of the frame where "via-inspector" stylesheet should be created.

    Dim params As New Dictionary
    params("frameId") = CStr(frameId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.createStyleSheet", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("styleSheetId") Then createStyleSheet = results("styleSheetId")
    End If
End Function

' Disables the CSS agent for the given page.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "CSS.disable", params
End Sub

' Enables the CSS agent for the given page. Clients should not assume that the CSS agent has been
' enabled until the result of this command is received.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "CSS.enable", params
End Sub

' Ensures that the given node will have specified pseudo-classes whenever its style is computed by
' the browser.
Public Sub forcePseudoState( _
    ByVal nodeId AS Long, _
    ByVal forcedPseudoClasses AS Collection _
)
    ' nodeId: DOM.NodeId The element id for which to force the pseudo state.
    ' forcedPseudoClasses: array Element pseudo classes to force when computing the element's style.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    Set params("forcedPseudoClasses") = forcedPseudoClasses

    cdp.InvokeMethod "CSS.forcePseudoState", params
End Sub

' 
' Returns:  - 
Public Function getBackgroundColors( _
    ByVal nodeId AS Long _
) AS Dictionary
    ' nodeId: DOM.NodeId Id of the node to get background colors for.

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.getBackgroundColors", params)
    If cdp.ErrorCode = 0 Then
        Set getBackgroundColors = results
    End If
End Function

' Returns the computed style for a DOM node identified by `nodeId`.
' Returns: computedStyle - Computed style for the specified DOM node.
Public Function getComputedStyleForNode( _
    ByVal nodeId AS Long _
) AS Collection
    ' nodeId: DOM.NodeId 

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.getComputedStyleForNode", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("computedStyle") Then Set getComputedStyleForNode = results("computedStyle")
    End If
End Function

' Returns the styles defined inline (explicitly in the "style" attribute and implicitly, using DOM
' attributes) for a DOM node identified by `nodeId`.
' Returns:  - 
Public Function getInlineStylesForNode( _
    ByVal nodeId AS Long _
) AS Dictionary
    ' nodeId: DOM.NodeId 

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.getInlineStylesForNode", params)
    If cdp.ErrorCode = 0 Then
        Set getInlineStylesForNode = results
    End If
End Function

' Returns requested styles for a DOM node identified by `nodeId`.
' Returns:  - 
Public Function getMatchedStylesForNode( _
    ByVal nodeId AS Long _
) AS Dictionary
    ' nodeId: DOM.NodeId 

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.getMatchedStylesForNode", params)
    If cdp.ErrorCode = 0 Then
        Set getMatchedStylesForNode = results
    End If
End Function

' Returns all media queries parsed by the rendering engine.
' Returns: medias - 
Public Function getMediaQueries() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.getMediaQueries", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("medias") Then Set getMediaQueries = results("medias")
    End If
End Function

' Requests information about platform fonts which we used to render child TextNodes in the given
' node.
' Returns: fonts - Usage statistics for every employed platform font.
Public Function getPlatformFontsForNode( _
    ByVal nodeId AS Long _
) AS Collection
    ' nodeId: DOM.NodeId 

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.getPlatformFontsForNode", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("fonts") Then Set getPlatformFontsForNode = results("fonts")
    End If
End Function

' Returns the current textual content for a stylesheet.
' Returns: text - The stylesheet text.
Public Function getStyleSheetText( _
    ByVal styleSheetId AS string _
) AS string
    ' styleSheetId: StyleSheetId 

    Dim params As New Dictionary
    params("styleSheetId") = CStr(styleSheetId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.getStyleSheetText", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("text") Then getStyleSheetText = results("text")
    End If
End Function

' Starts tracking the given computed styles for updates. The specified array of properties
' replaces the one previously specified. Pass empty array to disable tracking.
' Use takeComputedStyleUpdates to retrieve the list of nodes that had properties modified.
' The changes to computed style properties are only tracked for nodes pushed to the front-end
' by the DOM agent. If no changes to the tracked properties occur after the node has been pushed
' to the front-end, no updates will be issued for the node.
' Experimental
Public Sub trackComputedStyleUpdates( _
    ByVal propertiesToTrack AS Collection _
)
    ' propertiesToTrack: array 

    Dim params As New Dictionary
    Set params("propertiesToTrack") = propertiesToTrack

    cdp.InvokeMethod "CSS.trackComputedStyleUpdates", params
End Sub

' Polls the next batch of computed style updates.
' Returns: nodeIds - The list of node Ids that have their tracked computed styles updated
' Experimental
Public Function takeComputedStyleUpdates() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.takeComputedStyleUpdates", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodeIds") Then Set takeComputedStyleUpdates = results("nodeIds")
    End If
End Function

' Find a rule with the given active property for the given node and set the new value for this
' property
Public Sub setEffectivePropertyValueForNode( _
    ByVal nodeId AS Long, _
    ByVal propertyName AS string, _
    ByVal value AS string _
)
    ' nodeId: DOM.NodeId The element id for which to set property.
    ' propertyName: string 
    ' value: string 

    Dim params As New Dictionary
    params("nodeId") = CLng(nodeId)
    params("propertyName") = CStr(propertyName)
    params("value") = CStr(value)

    cdp.InvokeMethod "CSS.setEffectivePropertyValueForNode", params
End Sub

' Modifies the keyframe rule key text.
' Returns: keyText - The resulting key text after modification.
Public Function setKeyframeKey( _
    ByVal styleSheetId AS string, _
    ByVal range AS Dictionary, _
    ByVal keyText AS string _
) AS Dictionary
    ' styleSheetId: StyleSheetId 
    ' range: SourceRange 
    ' keyText: string 

    Dim params As New Dictionary
    params("styleSheetId") = CStr(styleSheetId)
    Set params("range") = range
    params("keyText") = CStr(keyText)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.setKeyframeKey", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("keyText") Then Set setKeyframeKey = results("keyText")
    End If
End Function

' Modifies the rule selector.
' Returns: media - The resulting CSS media rule after modification.
Public Function setMediaText( _
    ByVal styleSheetId AS string, _
    ByVal range AS Dictionary, _
    ByVal text AS string _
) AS Dictionary
    ' styleSheetId: StyleSheetId 
    ' range: SourceRange 
    ' text: string 

    Dim params As New Dictionary
    params("styleSheetId") = CStr(styleSheetId)
    Set params("range") = range
    params("text") = CStr(text)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.setMediaText", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("media") Then Set setMediaText = results("media")
    End If
End Function

' Modifies the expression of a container query.
' Returns: containerQuery - The resulting CSS container query rule after modification.
' Experimental
Public Function setContainerQueryText( _
    ByVal styleSheetId AS string, _
    ByVal range AS Dictionary, _
    ByVal text AS string _
) AS Dictionary
    ' styleSheetId: StyleSheetId 
    ' range: SourceRange 
    ' text: string 

    Dim params As New Dictionary
    params("styleSheetId") = CStr(styleSheetId)
    Set params("range") = range
    params("text") = CStr(text)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.setContainerQueryText", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("containerQuery") Then Set setContainerQueryText = results("containerQuery")
    End If
End Function

' Modifies the rule selector.
' Returns: selectorList - The resulting selector list after modification.
Public Function setRuleSelector( _
    ByVal styleSheetId AS string, _
    ByVal range AS Dictionary, _
    ByVal selector AS string _
) AS Dictionary
    ' styleSheetId: StyleSheetId 
    ' range: SourceRange 
    ' selector: string 

    Dim params As New Dictionary
    params("styleSheetId") = CStr(styleSheetId)
    Set params("range") = range
    params("selector") = CStr(selector)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.setRuleSelector", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("selectorList") Then Set setRuleSelector = results("selectorList")
    End If
End Function

' Sets the new stylesheet text.
' Returns: sourceMapURL - URL of source map associated with script (if any).
Public Function setStyleSheetText( _
    ByVal styleSheetId AS string, _
    ByVal text AS string _
) AS string
    ' styleSheetId: StyleSheetId 
    ' text: string 

    Dim params As New Dictionary
    params("styleSheetId") = CStr(styleSheetId)
    params("text") = CStr(text)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.setStyleSheetText", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("sourceMapURL") Then setStyleSheetText = results("sourceMapURL")
    End If
End Function

' Applies specified style edits one after another in the given order.
' Returns: styles - The resulting styles after modification.
Public Function setStyleTexts( _
    ByVal edits AS Collection _
) AS Collection
    ' edits: array 

    Dim params As New Dictionary
    Set params("edits") = edits

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.setStyleTexts", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("styles") Then Set setStyleTexts = results("styles")
    End If
End Function

' Enables the selector recording.
Public Sub startRuleUsageTracking()
    Dim params As New Dictionary
    cdp.InvokeMethod "CSS.startRuleUsageTracking", params
End Sub

' Stop tracking rule usage and return the list of rules that were used since last call to
' `takeCoverageDelta` (or since start of coverage instrumentation)
' Returns: ruleUsage - 
Public Function stopRuleUsageTracking() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.stopRuleUsageTracking", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("ruleUsage") Then Set stopRuleUsageTracking = results("ruleUsage")
    End If
End Function

' Obtain list of rules that became used since last call to this method (or since start of coverage
' instrumentation)
' Returns:  - 
Public Function takeCoverageDelta() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("CSS.takeCoverageDelta", params)
    If cdp.ErrorCode = 0 Then
        Set takeCoverageDelta = results
    End If
End Function

' Enables/disables rendering of local CSS fonts (enabled by default).
' Experimental
Public Sub setLocalFontsEnabled( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean Whether rendering of local fonts is enabled.

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "CSS.setLocalFontsEnabled", params
End Sub

