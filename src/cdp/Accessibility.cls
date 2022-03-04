VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAccessibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpAccessibility"
' Accessibility [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Unique accessibility node identifier.
' AXNodeId AS string

' Enum of possible property types.
' AXValueType AS string
Public Enum AXValueType
    AXVT_boolean
    AXVT_tristate
    AXVT_booleanOrUndefined
    AXVT_idref
    AXVT_idrefList
    AXVT_integer
    AXVT_node
    AXVT_nodeList
    AXVT_number
    AXVT_string
    AXVT_computedString
    AXVT_token
    AXVT_tokenList
    AXVT_domRelation
    AXVT_role
    AXVT_internalRole
    AXVT_valueUndefined
End Enum

' Enum of possible property sources.
' AXValueSourceType AS string
Public Enum AXValueSourceType
    AXVST_attribute
    AXVST_implicit
    AXVST_style
    AXVST_contents
    AXVST_placeholder
    AXVST_relatedElement
End Enum

' Enum of possible native property sources (as a subtype of a particular AXValueSourceType).
' AXValueNativeSourceType AS string
Public Enum AXValueNativeSourceType
    AXVNST_figcaption
    AXVNST_label
    AXVNST_labelfor
    AXVNST_labelwrapped
    AXVNST_legend
    AXVNST_rubyannotation
    AXVNST_tablecaption
    AXVNST_title
    AXVNST_other
End Enum

' A single source for a computed AX property.
' AXValueSource AS object

' 
' AXRelatedNode AS object

' 
' AXProperty AS object

' A single computed AX property.
' AXValue AS object

' Values of AXProperty name:
'   - from 'busy' to 'roledescription': states which apply to every AX node
'   - from 'live' to 'root': attributes which apply to nodes in live regions
'   - from 'autocomplete' to 'valuetext': attributes which apply to widgets
'   - from 'checked' to 'selected': states which apply to widgets
'   - from 'activedescendant' to 'owns' - relationships between elements other than parent/child/sibling.
' AXPropertyName AS string
Public Enum AXPropertyName
    AXPN_busy
    AXPN_disabled
    AXPN_editable
    AXPN_focusable
    AXPN_focused
    AXPN_hidden
    AXPN_hiddenRoot
    AXPN_invalid
    AXPN_keyshortcuts
    AXPN_settable
    AXPN_roledescription
    AXPN_live
    AXPN_atomic
    AXPN_relevant
    AXPN_root
    AXPN_autocomplete
    AXPN_hasPopup
    AXPN_level
    AXPN_multiselectable
    AXPN_orientation
    AXPN_multiline
    AXPN_readonly
    AXPN_required
    AXPN_valuemin
    AXPN_valuemax
    AXPN_valuetext
    AXPN_checked
    AXPN_expanded
    AXPN_modal
    AXPN_pressed
    AXPN_selected
    AXPN_activedescendant
    AXPN_controls
    AXPN_describedby
    AXPN_details
    AXPN_errormessage
    AXPN_flowto
    AXPN_labelledby
    AXPN_owns
End Enum

' A node in the accessibility tree.
' AXNode AS object


Public Function AXValueTypeToString(ByVal axvt As AXValueType) As String
    Dim retVal As String
    Select Case axvt
        Case AXVT_boolean
            retVal = "boolean"
        Case AXVT_tristate
            retVal = "tristate"
        Case AXVT_booleanOrUndefined
            retVal = "booleanOrUndefined"
        Case AXVT_idref
            retVal = "idref"
        Case AXVT_idrefList
            retVal = "idrefList"
        Case AXVT_integer
            retVal = "integer"
        Case AXVT_node
            retVal = "node"
        Case AXVT_nodeList
            retVal = "nodeList"
        Case AXVT_number
            retVal = "number"
        Case AXVT_string
            retVal = "string"
        Case AXVT_computedString
            retVal = "computedString"
        Case AXVT_token
            retVal = "token"
        Case AXVT_tokenList
            retVal = "tokenList"
        Case AXVT_domRelation
            retVal = "domRelation"
        Case AXVT_role
            retVal = "role"
        Case AXVT_internalRole
            retVal = "internalRole"
        Case AXVT_valueUndefined
            retVal = "valueUndefined"
        Case Else
            Debug.Print "Warning, unknown value " & axvt
    End Select
    AXValueTypeToString = retVal
End Function

Public Function StringToAXValueType(ByVal s As String) As AXValueType
    Dim retVal As AXValueType
    Select Case s
        Case "boolean"
            retVal = AXVT_boolean
        Case "tristate"
            retVal = AXVT_tristate
        Case "booleanOrUndefined"
            retVal = AXVT_booleanOrUndefined
        Case "idref"
            retVal = AXVT_idref
        Case "idrefList"
            retVal = AXVT_idrefList
        Case "integer"
            retVal = AXVT_integer
        Case "node"
            retVal = AXVT_node
        Case "nodeList"
            retVal = AXVT_nodeList
        Case "number"
            retVal = AXVT_number
        Case "string"
            retVal = AXVT_string
        Case "computedString"
            retVal = AXVT_computedString
        Case "token"
            retVal = AXVT_token
        Case "tokenList"
            retVal = AXVT_tokenList
        Case "domRelation"
            retVal = AXVT_domRelation
        Case "role"
            retVal = AXVT_role
        Case "internalRole"
            retVal = AXVT_internalRole
        Case "valueUndefined"
            retVal = AXVT_valueUndefined
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAXValueType = retVal
End Function


Public Function AXValueSourceTypeToString(ByVal axvst As AXValueSourceType) As String
    Dim retVal As String
    Select Case axvst
        Case AXVST_attribute
            retVal = "attribute"
        Case AXVST_implicit
            retVal = "implicit"
        Case AXVST_style
            retVal = "style"
        Case AXVST_contents
            retVal = "contents"
        Case AXVST_placeholder
            retVal = "placeholder"
        Case AXVST_relatedElement
            retVal = "relatedElement"
        Case Else
            Debug.Print "Warning, unknown value " & axvst
    End Select
    AXValueSourceTypeToString = retVal
End Function

Public Function StringToAXValueSourceType(ByVal s As String) As AXValueSourceType
    Dim retVal As AXValueSourceType
    Select Case s
        Case "attribute"
            retVal = AXVST_attribute
        Case "implicit"
            retVal = AXVST_implicit
        Case "style"
            retVal = AXVST_style
        Case "contents"
            retVal = AXVST_contents
        Case "placeholder"
            retVal = AXVST_placeholder
        Case "relatedElement"
            retVal = AXVST_relatedElement
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAXValueSourceType = retVal
End Function


Public Function AXValueNativeSourceTypeToString(ByVal axvnst As AXValueNativeSourceType) As String
    Dim retVal As String
    Select Case axvnst
        Case AXVNST_figcaption
            retVal = "figcaption"
        Case AXVNST_label
            retVal = "label"
        Case AXVNST_labelfor
            retVal = "labelfor"
        Case AXVNST_labelwrapped
            retVal = "labelwrapped"
        Case AXVNST_legend
            retVal = "legend"
        Case AXVNST_rubyannotation
            retVal = "rubyannotation"
        Case AXVNST_tablecaption
            retVal = "tablecaption"
        Case AXVNST_title
            retVal = "title"
        Case AXVNST_other
            retVal = "other"
        Case Else
            Debug.Print "Warning, unknown value " & axvnst
    End Select
    AXValueNativeSourceTypeToString = retVal
End Function

Public Function StringToAXValueNativeSourceType(ByVal s As String) As AXValueNativeSourceType
    Dim retVal As AXValueNativeSourceType
    Select Case s
        Case "figcaption"
            retVal = AXVNST_figcaption
        Case "label"
            retVal = AXVNST_label
        Case "labelfor"
            retVal = AXVNST_labelfor
        Case "labelwrapped"
            retVal = AXVNST_labelwrapped
        Case "legend"
            retVal = AXVNST_legend
        Case "rubyannotation"
            retVal = AXVNST_rubyannotation
        Case "tablecaption"
            retVal = AXVNST_tablecaption
        Case "title"
            retVal = AXVNST_title
        Case "other"
            retVal = AXVNST_other
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAXValueNativeSourceType = retVal
End Function


Public Function AXPropertyNameToString(ByVal axpn As AXPropertyName) As String
    Dim retVal As String
    Select Case axpn
        Case AXPN_busy
            retVal = "busy"
        Case AXPN_disabled
            retVal = "disabled"
        Case AXPN_editable
            retVal = "editable"
        Case AXPN_focusable
            retVal = "focusable"
        Case AXPN_focused
            retVal = "focused"
        Case AXPN_hidden
            retVal = "hidden"
        Case AXPN_hiddenRoot
            retVal = "hiddenRoot"
        Case AXPN_invalid
            retVal = "invalid"
        Case AXPN_keyshortcuts
            retVal = "keyshortcuts"
        Case AXPN_settable
            retVal = "settable"
        Case AXPN_roledescription
            retVal = "roledescription"
        Case AXPN_live
            retVal = "live"
        Case AXPN_atomic
            retVal = "atomic"
        Case AXPN_relevant
            retVal = "relevant"
        Case AXPN_root
            retVal = "root"
        Case AXPN_autocomplete
            retVal = "autocomplete"
        Case AXPN_hasPopup
            retVal = "hasPopup"
        Case AXPN_level
            retVal = "level"
        Case AXPN_multiselectable
            retVal = "multiselectable"
        Case AXPN_orientation
            retVal = "orientation"
        Case AXPN_multiline
            retVal = "multiline"
        Case AXPN_readonly
            retVal = "readonly"
        Case AXPN_required
            retVal = "required"
        Case AXPN_valuemin
            retVal = "valuemin"
        Case AXPN_valuemax
            retVal = "valuemax"
        Case AXPN_valuetext
            retVal = "valuetext"
        Case AXPN_checked
            retVal = "checked"
        Case AXPN_expanded
            retVal = "expanded"
        Case AXPN_modal
            retVal = "modal"
        Case AXPN_pressed
            retVal = "pressed"
        Case AXPN_selected
            retVal = "selected"
        Case AXPN_activedescendant
            retVal = "activedescendant"
        Case AXPN_controls
            retVal = "controls"
        Case AXPN_describedby
            retVal = "describedby"
        Case AXPN_details
            retVal = "details"
        Case AXPN_errormessage
            retVal = "errormessage"
        Case AXPN_flowto
            retVal = "flowto"
        Case AXPN_labelledby
            retVal = "labelledby"
        Case AXPN_owns
            retVal = "owns"
        Case Else
            Debug.Print "Warning, unknown value " & axpn
    End Select
    AXPropertyNameToString = retVal
End Function

Public Function StringToAXPropertyName(ByVal s As String) As AXPropertyName
    Dim retVal As AXPropertyName
    Select Case s
        Case "busy"
            retVal = AXPN_busy
        Case "disabled"
            retVal = AXPN_disabled
        Case "editable"
            retVal = AXPN_editable
        Case "focusable"
            retVal = AXPN_focusable
        Case "focused"
            retVal = AXPN_focused
        Case "hidden"
            retVal = AXPN_hidden
        Case "hiddenRoot"
            retVal = AXPN_hiddenRoot
        Case "invalid"
            retVal = AXPN_invalid
        Case "keyshortcuts"
            retVal = AXPN_keyshortcuts
        Case "settable"
            retVal = AXPN_settable
        Case "roledescription"
            retVal = AXPN_roledescription
        Case "live"
            retVal = AXPN_live
        Case "atomic"
            retVal = AXPN_atomic
        Case "relevant"
            retVal = AXPN_relevant
        Case "root"
            retVal = AXPN_root
        Case "autocomplete"
            retVal = AXPN_autocomplete
        Case "hasPopup"
            retVal = AXPN_hasPopup
        Case "level"
            retVal = AXPN_level
        Case "multiselectable"
            retVal = AXPN_multiselectable
        Case "orientation"
            retVal = AXPN_orientation
        Case "multiline"
            retVal = AXPN_multiline
        Case "readonly"
            retVal = AXPN_readonly
        Case "required"
            retVal = AXPN_required
        Case "valuemin"
            retVal = AXPN_valuemin
        Case "valuemax"
            retVal = AXPN_valuemax
        Case "valuetext"
            retVal = AXPN_valuetext
        Case "checked"
            retVal = AXPN_checked
        Case "expanded"
            retVal = AXPN_expanded
        Case "modal"
            retVal = AXPN_modal
        Case "pressed"
            retVal = AXPN_pressed
        Case "selected"
            retVal = AXPN_selected
        Case "activedescendant"
            retVal = AXPN_activedescendant
        Case "controls"
            retVal = AXPN_controls
        Case "describedby"
            retVal = AXPN_describedby
        Case "details"
            retVal = AXPN_details
        Case "errormessage"
            retVal = AXPN_errormessage
        Case "flowto"
            retVal = AXPN_flowto
        Case "labelledby"
            retVal = AXPN_labelledby
        Case "owns"
            retVal = AXPN_owns
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAXPropertyName = retVal
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

' Disables the accessibility domain.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Accessibility.disable", params
End Sub

' Enables the accessibility domain which causes `AXNodeId`s to remain consistent between method calls.
' This turns on accessibility for the page, which can impact performance until accessibility is disabled.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Accessibility.enable", params
End Sub

' Fetches the accessibility node and partial accessibility tree for this DOM node, if it exists.
' Returns: nodes - The `Accessibility.AXNode` for this DOM node, if it exists, plus its ancestors, siblings and
 '   children, if requested.
' Experimental
Public Function getPartialAXTree( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant, _
    Optional ByVal fetchRelatives AS Variant _
) AS Collection
    ' nodeId: DOM.NodeId(optional) Identifier of the node to get the partial accessibility tree for.
    ' backendNodeId: DOM.BackendNodeId(optional) Identifier of the backend node to get the partial accessibility tree for.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper to get the partial accessibility tree for.
    ' fetchRelatives: boolean(optional) Whether to fetch this nodes ancestors, siblings and children. Defaults to true.

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)
    If Not IsMissing(fetchRelatives) Then params("fetchRelatives") = CBool(fetchRelatives)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Accessibility.getPartialAXTree", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodes") Then Set getPartialAXTree = results("nodes")
    End If
End Function

' Fetches the entire accessibility tree for the root Document
' Returns: nodes - 
' Experimental
Public Function getFullAXTree( _
    Optional ByVal depth AS Variant, _
    Optional ByVal max_depth AS Variant, _
    Optional ByVal frameId AS Variant _
) AS Collection
    ' depth: integer(optional) The maximum depth at which descendants of the root node should be retrieved.
    '   If omitted, the full tree is returned.
    ' max_depth: integer(optional) Deprecated. This parameter has been renamed to `depth`. If depth is not provided, max_depth will be used.
    ' frameId: Page.FrameId(optional) The frame for whose document the AX tree should be retrieved.
    '   If omited, the root frame is used.

    Dim params As New Dictionary
    If Not IsMissing(depth) Then params("depth") = CLng(depth)
    If Not IsMissing(max_depth) Then params("max_depth") = CLng(max_depth)
    If Not IsMissing(frameId) Then params("frameId") = CStr(frameId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Accessibility.getFullAXTree", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodes") Then Set getFullAXTree = results("nodes")
    End If
End Function

' Fetches a particular accessibility node by AXNodeId.
' Requires `enable()` to have been called previously.
' Returns: nodes - 
' Experimental
Public Function getChildAXNodes( _
    ByVal id AS string, _
    Optional ByVal frameId AS Variant _
) AS Collection
    ' id: AXNodeId 
    ' frameId: Page.FrameId(optional) The frame in whose document the node resides.
    '   If omitted, the root frame is used.

    Dim params As New Dictionary
    params("id") = CStr(id)
    If Not IsMissing(frameId) Then params("frameId") = CStr(frameId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Accessibility.getChildAXNodes", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodes") Then Set getChildAXNodes = results("nodes")
    End If
End Function

' Query a DOM node's accessibility subtree for accessible name and role.
' This command computes the name and role for all nodes in the subtree, including those that are
' ignored for accessibility, and returns those that mactch the specified name and role. If no DOM
' node is specified, or the DOM node does not exist, the command returns an error. If neither
' `accessibleName` or `role` is specified, it returns all the accessibility nodes in the subtree.
' Returns: nodes - A list of `Accessibility.AXNode` matching the specified attributes,
 '   including nodes that are ignored for accessibility.
' Experimental
Public Function queryAXTree( _
    Optional ByVal nodeId AS Variant, _
    Optional ByVal backendNodeId AS Variant, _
    Optional ByVal objectId AS Variant, _
    Optional ByVal accessibleName AS Variant, _
    Optional ByVal role AS Variant _
) AS Collection
    ' nodeId: DOM.NodeId(optional) Identifier of the node for the root to query.
    ' backendNodeId: DOM.BackendNodeId(optional) Identifier of the backend node for the root to query.
    ' objectId: Runtime.RemoteObjectId(optional) JavaScript object id of the node wrapper for the root to query.
    ' accessibleName: string(optional) Find nodes with this computed name.
    ' role: string(optional) Find nodes with this computed role.

    Dim params As New Dictionary
    If Not IsMissing(nodeId) Then params("nodeId") = CLng(nodeId)
    If Not IsMissing(backendNodeId) Then params("backendNodeId") = CLng(backendNodeId)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)
    If Not IsMissing(accessibleName) Then params("accessibleName") = CStr(accessibleName)
    If Not IsMissing(role) Then params("role") = CStr(role)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Accessibility.queryAXTree", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("nodes") Then Set queryAXTree = results("nodes")
    End If
End Function

