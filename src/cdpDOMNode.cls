VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Node"
' DOM.Node
' DOM interaction is implemented in terms of mirror objects that represent the actual DOM nodes.
'   DOMNode is a base node mirror type.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' to allow convenience form of calling node.fn(...) instead of browser.fn(node,...)
Private Browser As AutomateBrowser
' resolve javascript reference to node, only done if browser is provided
Public resolveNode As cdpRuntimeRemoteObject
' cache object properties, initialized on demand not during init as takes time to serialize and deserialize to do automatically
Private objectProperties As Dictionary ' of cdpRuntimePropertyDescriptor
' cache of parent node, to allow traversing next/prev sibling
Private parentNode As cdpDOMNode

' which method to use to set a node's "value"
Public Enum SetNodeValueMode
    SNV_Automatic ' guess best method to set value, may not work correctly
    SNV_Attribute
    SNV_Value
    SNV_NodeValue
    SNV_FakeInput
    SNV_Clipboard
End Enum

Private autoSetMode As SetNodeValueMode

' modifiers to simulating key input
Private Enum KbModifier
    None = 0
    Alt = 1
    Ctrl = 2
    MetaCommand = 4
    Shift = 8
End Enum

' valid node types
Public Enum NodeTypeValue
    NT_ELEMENT = 1              ' An element node such as <p> or <div>
    NT_ATTRIBUTE                ' An attribute of an element
    NT_TEXT                     ' The textual content (actual "text") inside an element or attribute
    NT_CDATA_SECTION            ' A CDATASection, such as <!CDATA[[ … ]]>
    NT_ENTITY_REFERENCE         ' deprecated / unused
    NT_ENTITY                   ' deprecated / unused
    NT_PROCESSING_INSTRUCTION   ' A processingInstruction of an XML document, such as <?xml-stylesheet … ?>
    NT_COMMENT                  ' A comment node, such as <!-- … -->
    NT_DOCUMENT                 ' A document node, the root node of the DOM tree
    NT_DOCUMENT_TYPE            ' A documentType node, such as <!DOCTYPE html>
    NT_DOCUMENT_FRAGMENT        ' A documentFragment node
    NT_NOTATION                 ' deprecated / unused
End Enum

' nodeId : NodeId
' Node identifier that is passed into the rest of the DOM messages as the `nodeId`. Backend
    '   will only push node with given `id` once. It is aware of all requested nodes and will only
    '   fire DOM events for nodes known to the client.
Public nodeId As Long

' parentId : NodeId
' Optional
' The id of the parent node if any.
Public parentId As Long

' backendNodeId : BackendNodeId
' The BackendNodeId for this node.
Public backendNodeId As Long

' nodeType : integer
' `Node`'s nodeType.
Public nodeType As Long

' nodeName : string
' `Node`'s nodeName.
Public nodeName As String

' localName : string
' `Node`'s localName.
Public localName As String

' nodeValue : string
' `Node`'s nodeValue.
Public nodeValue As String

' childNodeCount : integer
' Optional
' Child count for `Container` nodes.
Public childNodeCount As Long

' children : array of cdpDOMNode
' Optional
' Child nodes of this node when requested with children.
Public children As Collection

' attributes : array of string
' Optional
' Attributes of the `Element` node in the form of flat array `[name1, value1, name2, value2]`.
Public attributes As Collection

' documentURL : string
' Optional
' Document URL that `Document` or `FrameOwner` node points to.
Public documentURL As String

' baseURL : string
' Optional
' Base URL that `Document` or `FrameOwner` node uses for URL completion.
Public baseURL As String

' publicId : string
' Optional
' `DocumentType`'s publicId.
Public publicId As String

' systemId : string
' Optional
' `DocumentType`'s systemId.
Public systemId As String

' internalSubset : string
' Optional
' `DocumentType`'s internalSubset.
Public internalSubset As String

' xmlVersion : string
' Optional
' `Document`'s XML version in case of XML documents.
Public xmlVersion As String

' name : string
' Optional
' `Attr`'s name.
Public name As String

' value : string
' Optional
' `Attr`'s value.
Public value As String

' pseudoType : PseudoType
' Optional
' Pseudo element type for this node.
Public pseudoType As String

' shadowRootType : ShadowRootType
' Optional
' Shadow root type.
Public shadowRootType As String

' frameId : Page.FrameId
' Optional
' Frame ID for frame owner elements.
Public frameId As String

' contentDocument : Node
' Optional
' Content document for frame owner elements.
Public contentDocument As cdpDOMNode

' shadowRoots : array of cdpDOMNode
' Optional
' Shadow root list for given element host.
Public shadowRoots As Collection

' templateContent : Node
' Optional
' Content document fragment for template elements.
Public templateContent As cdpDOMNode

' pseudoElements : array of cdpDOMNode
' Optional
' Pseudo elements associated with this node.
Public pseudoElements As Collection

' importedDocument : Node
' Optional
' Deprecated, as the HTML Imports API has been removed (crbug.com/937746).
    '   This property used to return the imported document for the HTMLImport links.
    '   The property is always undefined now.
Public importedDocument As cdpDOMNode

' distributedNodes : array of cdpDOMBackendNode
' Optional
' Distributed nodes for given insertion point.
Public distributedNodes As Collection

' isSVG : boolean
' Optional
' Whether the node is SVG.
Public isSVG As Variant ' boolean

' compatibilityMode : CompatibilityMode
' Optional
'
Public compatibilityMode As String

Public Function init(ByVal obj As Dictionary, Optional ByVal b As AutomateBrowser, Optional ByVal expectedNodeId As Integer, Optional ByVal parent As cdpDOMNode = Nothing) As cdpDOMNode
Attribute init.VB_Description = "Initialize class from Dictionary returned by CDP method."
    On Error GoTo ErrHandler
    Dim v As Variant

    Let nodeId = CLng(obj.Item("nodeId"))
    
    If obj.Exists("parentId") Then Let parentId = CLng(obj.Item("parentId"))
    Let backendNodeId = CLng(obj.Item("backendNodeId"))
    Let nodeType = CLng(obj.Item("nodeType"))
    Let nodeName = CStr(obj.Item("nodeName"))
    Let localName = CStr(obj.Item("localName"))
    Let nodeValue = CStr(obj.Item("nodeValue"))
    If obj.Exists("childNodeCount") Then Let childNodeCount = CLng(obj.Item("childNodeCount"))
    If obj.Exists("children") Then
        Set children = New Collection
        For Each v In obj.Item("children")
            Dim obj_children As cdpDOMNode: Set obj_children = New cdpDOMNode
            obj_children.init v, b, parent:=Me
            children.Add obj_children
            Set obj_children = Nothing
        Next v
    End If
    If obj.Exists("attributes") Then
        Set attributes = obj.Item("attributes")
        'For Each v In obj.Item("attributes")
        '    attributes.Add v
        'Next v
    End If
    If obj.Exists("documentURL") Then Let documentURL = CStr(obj.Item("documentURL"))
    If obj.Exists("baseURL") Then Let baseURL = CStr(obj.Item("baseURL"))
    If obj.Exists("publicId") Then Let publicId = CStr(obj.Item("publicId"))
    If obj.Exists("systemId") Then Let systemId = CStr(obj.Item("systemId"))
    If obj.Exists("internalSubset") Then Let internalSubset = CStr(obj.Item("internalSubset"))
    If obj.Exists("xmlVersion") Then Let xmlVersion = CStr(obj.Item("xmlVersion"))
    If obj.Exists("name") Then Let name = CStr(obj.Item("name"))
    If obj.Exists("value") Then Let value = CStr(obj.Item("value"))
    If obj.Exists("pseudoType") Then Let pseudoType = CStr(obj.Item("pseudoType"))
    If obj.Exists("shadowRootType") Then Let shadowRootType = CStr(obj.Item("shadowRootType"))
    If obj.Exists("frameId") Then Let frameId = CStr(obj.Item("frameId"))
    If obj.Exists("contentDocument") Then
        Set contentDocument = New cdpDOMNode
        contentDocument.init obj.Item("contentDocument"), b, parent:=Me
    End If
    If obj.Exists("shadowRoots") Then
        Set shadowRoots = New Collection
        For Each v In obj.Item("shadowRoots")
            Dim obj_shadowRoots As cdpDOMNode: Set obj_shadowRoots = New cdpDOMNode
            obj_shadowRoots.init v, b, parent:=Me
            shadowRoots.Add obj_shadowRoots
            Set obj_shadowRoots = Nothing
        Next v
    End If
    If obj.Exists("templateContent") Then
        Set templateContent = New cdpDOMNode
        templateContent.init obj.Item("templateContent"), b, parent:=Me
    End If
    If obj.Exists("pseudoElements") Then
        Set pseudoElements = New Collection
        For Each v In obj.Item("pseudoElements")
            Dim obj_pseudoElements As cdpDOMNode: Set obj_pseudoElements = New cdpDOMNode
            obj_pseudoElements.init v, b, parent:=Me
            pseudoElements.Add obj_pseudoElements
            Set obj_pseudoElements = Nothing
        Next v
    End If
    If obj.Exists("importedDocument") Then
        Set importedDocument = New cdpDOMNode
        importedDocument.init obj.Item("importedDocument"), b, parent:=Me
    End If
    If obj.Exists("distributedNodes") Then
        Set distributedNodes = New Collection
        For Each v In obj.Item("distributedNodes")
            Dim obj_distributedNodes As cdpDOMBackendNode: Set obj_distributedNodes = New cdpDOMBackendNode
            obj_distributedNodes.init v
            distributedNodes.Add obj_distributedNodes
            Set obj_distributedNodes = Nothing
        Next v
    End If
    If obj.Exists("isSVG") Then Let isSVG = CBool(obj.Item("isSVG"))
    If obj.Exists("compatibilityMode") Then Let compatibilityMode = CStr(obj.Item("compatibilityMode"))

    If nodeId = 0 And expectedNodeId <> 0 Then nodeId = expectedNodeId ' note: this gets stale quickly, but if we know it then don't forget it
    If Not parent Is Nothing Then
        Set parentNode = parent
        If parentId = 0 Then parentId = parent.nodeId
    End If
    If Not b Is Nothing Then
        Set Browser = b
        Set v = Nothing
        If nodeId > 0 Then
            Set v = Browser.cdp.DOM.resolveNode(nodeId)
        ElseIf backendNodeId > 0 Then
            Set v = Browser.cdp.DOM.resolveNode(backendNodeId:=backendNodeId)
        End If
        If Not v Is Nothing Then
            Set resolveNode = New cdpRuntimeRemoteObject
            resolveNode.init v
        End If
    End If

    Set init = Me
    Exit Function
ErrHandler:
    Stop
    Resume Next
End Function

Public Function asDictionary() As Dictionary
    Dim dict As Dictionary: Set dict = New Dictionary

    dict("nodeId") = nodeId
    dict("parentId") = parentId
    dict("backendNodeId") = backendNodeId
    dict("nodeType") = nodeType
    dict("nodeName") = nodeName
    dict("localName") = localName
    dict("nodeValue") = nodeValue
    dict("childNodeCount") = childNodeCount
    Set dict("children") = children
    Set dict("attributes") = attributes
    If documentURL <> vbNullString Then dict("documentURL") = documentURL
    If baseURL <> vbNullString Then dict("baseURL") = baseURL
    If publicId <> vbNullString Then dict("publicId") = publicId
    If systemId <> vbNullString Then dict("systemId") = systemId
    If internalSubset <> vbNullString Then dict("internalSubset") = internalSubset
    If xmlVersion <> vbNullString Then dict("xmlVersion") = xmlVersion
    If name <> vbNullString Then dict("name") = name
    If value <> vbNullString Then dict("value") = value
    If pseudoType <> vbNullString Then dict("pseudoType") = pseudoType
    If shadowRootType <> vbNullString Then dict("shadowRootType") = shadowRootType
    If frameId <> vbNullString Then dict("frameId") = frameId
    If Not contentDocument Is Nothing Then Set dict("contentDocument") = contentDocument.asDictionary()
    Set dict("shadowRoots") = shadowRoots
    If Not templateContent Is Nothing Then Set dict("templateContent") = templateContent.asDictionary()
    Set dict("pseudoElements") = pseudoElements
    If Not importedDocument Is Nothing Then Set dict("importedDocument") = importedDocument.asDictionary()
    Set dict("distributedNodes") = distributedNodes
    If Not IsEmpty(isSVG) Then dict("isSVG") = isSVG
    If compatibilityMode <> vbNullString Then dict("compatibilityMode") = compatibilityMode

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set children = New Collection
    Set attributes = New Collection
    Set shadowRoots = New Collection
    Set pseudoElements = New Collection
    Set distributedNodes = New Collection
    
    autoSetMode = SNV_Automatic
End Sub

Private Sub Class_Terminate()
    Set Browser = Nothing
    Set resolveNode = Nothing
    Set objectProperties = Nothing
    Set parentNode = Nothing
    
    Set children = Nothing
    Set attributes = Nothing
    Set shadowRoots = Nothing
    Set pseudoElements = Nothing
    Set distributedNodes = Nothing
End Sub

' ************* Convenience functions *********************
Public Sub Click(Optional ByVal strategy As NavigationStrategy = NavigationStrategy.None)
    Browser.Click nodeId:=nodeId, backendNodeId:=backendNodeId, strategy:=strategy
End Sub


Public Function QuerySelector(ByVal selector As String, Optional ByVal frameId As String = vbNullString, _
                              Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False, _
                              Optional waitUntilFound As Boolean = True, Optional waitUntilVisible As Boolean = True) As cdpDOMNode
    Set QuerySelector = Browser.QuerySelector(selector, nodeId, frameId, depth, pierce, waitUntilFound, waitUntilVisible)
End Function

Public Function querySelectorAll(ByVal selector As String, Optional ByVal rootId As Long = 0, Optional ByVal frameId As String = vbNullString, _
                                 Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False) As Dictionary ' of cdpDomNode
    Set querySelectorAll = Browser.querySelectorAll(selector, nodeId, frameId, depth, pierce)
End Function


Public Sub focus()
    Browser.focus nodeId:=nodeId, backendNodeId:=backendNodeId
End Sub

Public Sub getCenter(ByRef x As Double, ByRef y As Double)
    Browser.getNodeCenter x, y, nodeId:=nodeId, backendNodeId:=backendNodeId
End Sub


Public Property Get elementValue() As String
    Dim child As cdpDOMNode, v As Variant
    
    ' look for shadowDom for value
    If shadowRoots.count > 0 Then
        ' we assume textContent is our value
        Set child = shadowRoots(1)
        elementValue = child.textContent
    End If
    
    ' if we didn't find a value (or its "")
    If elementValue = vbNullString Then
        'is the value itself set? use it
        If value <> vbNullString Then
            elementValue = value
        ' try nodeValue, which should correspond with value attribute but directly accessible
        ElseIf nodeValue <> vbNullString Then
            elementValue = nodeValue
        ' not yet then see if there is a child (not shadow) element that is a #Text node, e.g. for textarea like controls
        ElseIf childNodeCount > 0 Then
            For Each v In children
                Set child = v
                If child.nodeType = 3 Then
                    elementValue = child.elementValue
                    Exit Property
                End If
            Next v
        End If
    End If
    
    ' if all else fails, just return value attribute's value
    If elementValue = vbNullString Then
        elementValue = elementAttribute("value")
    End If
End Property

Public Property Let elementValue(ByVal newValue As String)
    setValue newValue, SetNodeValueMode.SNV_Automatic
End Property

Public Sub setValue(ByVal newValue As String, Optional mode As SetNodeValueMode = SetNodeValueMode.SNV_Automatic)
    Dim child As cdpDOMNode, v As Variant
    
    ' first focus the element (its ok if this fails usually)
    If nodeId <> 0 Then
        Browser.cdp.DOM.focus nodeId:=nodeId
        Browser.cdp.DOM.scrollIntoViewIfNeeded nodeId:=nodeId
    Else
        Browser.cdp.DOM.focus backendNodeId:=backendNodeId
        Browser.cdp.DOM.scrollIntoViewIfNeeded backendNodeId:=backendNodeId
    End If
    
    ' select anything currently set - don't do this unless direct input has already been done, as it simulates direct input and that breaks setting the value explicitly
    'Browser.cdp.SimulateInput.dispatchKeyEvent "keyDown", modifiers:=KbModifier.Ctrl, Text:="a", windowsVirtualKeyCode:=65
    'Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp"
    
    ' guess proper mode if automatic selected
    If mode = SetNodeValueMode.SNV_Automatic Then
        If autoSetMode <> SetNodeValueMode.SNV_Automatic Then
            mode = autoSetMode
        ElseIf nodeName = "INPUT" Then
            ' if node is INPUT then assume value stored in attribute
            mode = SetNodeValueMode.SNV_Attribute
        ElseIf childNodeCount > 0 Then
            ' set first Text node found's node value, e.g. assume like a TEXTAREA
            mode = SetNodeValueMode.SNV_NodeValue
        Else
            ' ???
            mode = SetNodeValueMode.SNV_Attribute
        End If
    End If
    
    ' set the value using requested method
    Select Case mode
        Case SetNodeValueMode.SNV_Attribute
            elementAttribute("value") = newValue
        Case SetNodeValueMode.SNV_Value
            Browser.cdp.DOM.setNodeValue nodeId, newValue ' ???
            ' update cached value as well
            'value = newValue
        Case SetNodeValueMode.SNV_NodeValue
            If nodeType = 3 Then ' is this a text node?
                Browser.cdp.DOM.setNodeValue nodeId, newValue
                ' update cached value as well
                'nodeValue = newValue
            Else
                ' if has children see if has Text node (type=3) and directly set its value
                For Each v In children
                    Set child = v
                    If child.nodeType = 3 Then
                        Browser.cdp.DOM.setNodeValue child.nodeId, newValue
                        ' update cached value as well
                        'child.nodeValue = newValue
                        Exit For
                    End If
                Next v
            End If
            Browser.cdp.DOM.setNodeValue nodeId, newValue
        Case SetNodeValueMode.SNV_Clipboard
            autoSetMode = SetNodeValueMode.SNV_Clipboard
            'elementAttribute("value") = vbNullString
            ' put string in clipboard
            SetClipboard newValue
            ' select all current text so replaced with new value, via select all
            Browser.cdp.SimulateInput.dispatchKeyEvent "rawKeyDown", modifiers:=KbModifier.Ctrl, Text:="a", unmodifiedText:="a", code:="KeyA", key:="KeyA", windowsVirtualKeyCode:=65
            Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp", modifiers:=KbModifier.Ctrl, Text:="a", unmodifiedText:="a", code:="KeyA", key:="KeyA"
            ' issue Ctrl-V to paste
            Browser.cdp.SimulateInput.dispatchKeyEvent "keyDown", modifiers:=KbModifier.Ctrl, code:="KeyV", windowsVirtualKeyCode:=86
            Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp"
        Case SetNodeValueMode.SNV_FakeInput
            autoSetMode = SetNodeValueMode.SNV_FakeInput
            'elementAttribute("value") = vbNullString
            ' we need to clear node first
            ' select all current text so replaced with new value, via select all
            Browser.cdp.SimulateInput.dispatchKeyEvent "rawKeyDown", modifiers:=KbModifier.Ctrl, Text:="a", unmodifiedText:="a", code:="KeyA", key:="KeyA", windowsVirtualKeyCode:=65
            Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp", modifiers:=KbModifier.Ctrl, Text:="a", unmodifiedText:="a", code:="KeyA", key:="KeyA"
            ' then dispatch string as a set of fake key events
            Dim c As String, ndx As Long
            For ndx = 1 To Len(newValue)
                c = Mid$(newValue, ndx, 1)
                Browser.cdp.SimulateInput.dispatchKeyEvent "keyDown", Text:=c
                ' doing all three can cause duplicate characters input, char doesn't always work though, so use keyDown/keyUp
                'Browser.cdp.SimulateInput.dispatchKeyEvent "char", Text:=c
                Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp"
            Next ndx
    End Select
    
    ' to avoid stale data, we need to update all attributes, values, shadowdom, etc
    Dim dict As Dictionary
    If nodeId <> 0 Then
        Set dict = Browser.cdp.DOM.describeNode(nodeId:=nodeId, depth:=-1, pierce:=True)
    ElseIf backendNodeId <> 0 Then
        Set dict = Browser.cdp.DOM.describeNode(backendNodeId:=backendNodeId, depth:=-1, pierce:=True)
    End If
    If Browser.cdp.ErrorCode = 0 Then
        Me.init dict, Browser, nodeId, parentNode
    End If
End Sub


Public Property Get elementAttribute(ByVal name As String) As String
    elementAttribute = Browser.getNodeAttribute(Me, name)
End Property

Public Property Let elementAttribute(ByVal name As String, ByVal value As String)
    Browser.cdp.DOM.setAttributeValue nodeId, name, value
    
    ' update cached value
    If Browser.cdp.ErrorCode = 0 Then
        Dim v As Variant, col As Collection
        Set attributes = New Collection
        Set col = Browser.cdp.DOM.getAttributes(nodeId)
        For Each v In col
            attributes.Add v
        Next v
    End If
End Property


' clears our cache so next call to getProperty (or related) retrieves [updated] property values
Public Sub refreshProperties()
    Set objectProperties = Nothing
End Sub

' returns specified javascript property associated with this node
' see Runtime.getProperties
' Note: this call may take a noticable amount of time, especially when called on multiple nodes with many properties
Public Function getProperty(ByVal propertyName As String) As cdpRuntimePropertyDescriptor
    Dim index As Integer
    Dim objectProperty As Dictionary
    
    If objectProperties Is Nothing Then Set objectProperties = Browser.getProperties(Me)
    If Not objectProperties Is Nothing Then
        If objectProperties.Exists(propertyName) Then
            Dim prop As cdpRuntimePropertyDescriptor
            Set prop = objectProperties(propertyName)
            Set getProperty = prop
        End If
    End If
End Function

' returns value of specified javascript property associated with this node
' see Runtime.getProperties
' Note: this call may take a noticable amount of time, especially when called on multiple nodes with many properties
Public Function getPropertyValue(ByVal propertyName As String) As String
    Dim index As Integer
    Dim objectProperty As Dictionary
    
    If objectProperties Is Nothing Then Set objectProperties = Browser.getProperties(Me)
    If Not objectProperties Is Nothing Then
        If objectProperties.Exists(propertyName) Then
            Dim prop As cdpRuntimePropertyDescriptor
            Set prop = objectProperties(propertyName)
            If IsNull(prop.value.value) Then
                getPropertyValue = vbNullString
            Else
                getPropertyValue = CStr(prop.value.value)
            End If
        End If
    End If
End Function

Public Function getIdOrName() As String
    getIdOrName = Browser.getNodeIdOrName(Me)
End Function

Public Property Get id() As String
    id = getPropertyValue("id")
End Property

Public Property Get className() As String
    className = getPropertyValue("className")
End Property

Public Property Get title() As String
    title = getPropertyValue("title")
End Property

Public Property Get textContent() As String
    textContent = getPropertyValue("textContent")
End Property

Public Property Get tagName() As String
    tagName = getPropertyValue("tagName")
End Property

Public Property Get outerText() As String
    outerText = getPropertyValue("outerText")
End Property

Public Property Get InnerText() As String
    InnerText = getPropertyValue("innerText")
End Property

Public Property Get innerHTML() As String
    innerHTML = getPropertyValue("innerHTML")
End Property

Public Property Get outerHTML() As String
    'outerHTML = getPropertyValue("outerHTML")
    outerHTML = Browser.cdp.DOM.getOuterHTML(objectId:=resolveNode.objectId)
End Property


' returns True only if element meets both criteria:
'   has a positive width and height and
'   CSS display property is not "none" or "hidden"
Public Function isVisible() As Boolean
    'isVisible = False
    
    ' first check if has a positive width and height
    Dim x As Double, y As Double
    getCenter x, y
    If (Abs(x) > 0#) And (Abs(y) > 0#) Then
        ' now get computed CSS display value
        Dim style As cdpCSSCSSComputedStyleProperty
        Set style = getCSSStyle("display")
        If Not style Is Nothing Then
            If (style.value <> "none") And (style.value <> "hidden") Then
                isVisible = True
            End If
        End If
    End If
End Function


' retrieves current CSS computed style for specified property
Public Function getCSSStyle(ByVal name As String) As cdpCSSCSSComputedStyleProperty
    Dim styles As Collection
    Browser.cdp.CSS.enable ' must be enabled or call to computer style will fail
    Set styles = Browser.cdp.CSS.getComputedStyleForNode(nodeId)
    Browser.cdp.CSS.disable
    Dim v As Variant
    For Each v In styles
        If v("name") = name Then
            Dim style As cdpCSSCSSComputedStyleProperty
            Set style = New cdpCSSCSSComputedStyleProperty
            style.init v
            Set getCSSStyle = style
            Exit Function
        End If
    Next v
End Function


' returns neighboring child node of parent
' returns Nothing if no next neighbor or parent unknown
Public Function NextSibling() As cdpDOMNode
    If Not parentNode Is Nothing Then
        Dim v As Variant, lastNode As cdpDOMNode
        For Each v In parentNode.children
            Dim node As cdpDOMNode
            Set node = v
            ' was previous child me, if so then return this one
            If lastNode Is Me Then
                Set NextSibling = node
                Exit Function
            End If
            Set lastNode = node
        Next v
    End If
End Function

' returns neighboring child node of parent
' returns Nothing if no previous neighbor or parent unknown
Public Function PrevSibling() As cdpDOMNode
    If Not parentNode Is Nothing Then
        Dim v As Variant, lastNode As cdpDOMNode
        For Each v In parentNode.children
            Dim node As cdpDOMNode
            Set node = v
            ' is current child me, if so then return last one
            If node Is Me Then
                Set PrevSibling = lastNode
                Exit Function
            End If
            Set lastNode = node
        Next v
    End If
End Function
