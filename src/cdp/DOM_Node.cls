VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' DOM.Node
' DOM interaction is implemented in terms of mirror objects that represent the actual DOM nodes.
'   DOMNode is a base node mirror type.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' to allow convenience form of calling node.fn(...) instead of browser.fn(node,...)
Private browser As AutomateBrowser
Private resolveNode As Dictionary
Private objectProperties As Dictionary

' which method to use to set a node's "value"
Public Enum SetNodeValueMode
    SNV_AsAttribute
    SNV_AsValue
    SNV_AsNodeValue
    SNV_AsFakeInput
End Enum

' nodeId : NodeId
' Node identifier that is passed into the rest of the DOM messages as the `nodeId`. Backend
    '   will only push node with given `id` once. It is aware of all requested nodes and will only
    '   fire DOM events for nodes known to the client.
Public nodeId AS integer

' parentId : NodeId
' Optional
' The id of the parent node if any.
Public parentId AS integer

' backendNodeId : BackendNodeId
' The BackendNodeId for this node.
Public backendNodeId AS integer

' nodeType : integer
' `Node`'s nodeType.
Public nodeType AS integer

' nodeName : string
' `Node`'s nodeName.
Public nodeName AS string

' localName : string
' `Node`'s localName.
Public localName AS string

' nodeValue : string
' `Node`'s nodeValue.
Public nodeValue AS string

' childNodeCount : integer
' Optional
' Child count for `Container` nodes.
Public childNodeCount AS integer

' children : array of cdpDOMNode
' Optional
' Child nodes of this node when requested with children.
Public children AS Collection

' attributes : array of string
' Optional
' Attributes of the `Element` node in the form of flat array `[name1, value1, name2, value2]`.
Public attributes AS Collection

' documentURL : string
' Optional
' Document URL that `Document` or `FrameOwner` node points to.
Public documentURL AS string

' baseURL : string
' Optional
' Base URL that `Document` or `FrameOwner` node uses for URL completion.
Public baseURL AS string

' publicId : string
' Optional
' `DocumentType`'s publicId.
Public publicId AS string

' systemId : string
' Optional
' `DocumentType`'s systemId.
Public systemId AS string

' internalSubset : string
' Optional
' `DocumentType`'s internalSubset.
Public internalSubset AS string

' xmlVersion : string
' Optional
' `Document`'s XML version in case of XML documents.
Public xmlVersion AS string

' name : string
' Optional
' `Attr`'s name.
Public name AS string

' value : string
' Optional
' `Attr`'s value.
Public value AS string

' pseudoType : PseudoType
' Optional
' Pseudo element type for this node.
Public pseudoType AS string

' shadowRootType : ShadowRootType
' Optional
' Shadow root type.
Public shadowRootType AS string

' frameId : Page.FrameId
' Optional
' Frame ID for frame owner elements.
Public frameId AS string

' contentDocument : Node
' Optional
' Content document for frame owner elements.
Public contentDocument AS cdpDOMNode

' shadowRoots : array of cdpDOMNode
' Optional
' Shadow root list for given element host.
Public shadowRoots AS Collection

' templateContent : Node
' Optional
' Content document fragment for template elements.
Public templateContent AS cdpDOMNode

' pseudoElements : array of cdpDOMNode
' Optional
' Pseudo elements associated with this node.
Public pseudoElements AS Collection

' importedDocument : Node
' Optional
' Deprecated, as the HTML Imports API has been removed (crbug.com/937746).
    '   This property used to return the imported document for the HTMLImport links.
    '   The property is always undefined now.
Public importedDocument AS cdpDOMNode

' distributedNodes : array of cdpDOMBackendNode
' Optional
' Distributed nodes for given insertion point.
Public distributedNodes AS Collection

' isSVG : boolean
' Optional
' Whether the node is SVG.
Public isSVG AS Variant ' boolean

' compatibilityMode : CompatibilityMode
' Optional
' 
Public compatibilityMode AS string

Public Function init(ByVal obj as Dictionary, Optional ByVal b As AutomateBrowser) As cdpDOMNode
    Dim v as Variant

    Let nodeId = CInt(obj.Item("nodeId"))
    If obj.Exists("parentId") Then Let parentId = CInt(obj.Item("parentId"))
    Let backendNodeId = CInt(obj.Item("backendNodeId"))
    Let nodeType = CInt(obj.Item("nodeType"))
    Let nodeName = CStr(obj.Item("nodeName"))
    Let localName = CStr(obj.Item("localName"))
    Let nodeValue = CStr(obj.Item("nodeValue"))
    If obj.Exists("childNodeCount") Then Let childNodeCount = CInt(obj.Item("childNodeCount"))
    If obj.Exists("children") Then
        For each v in obj.Item("children")
            Dim obj_children As cdpDOMNode: Set obj_children = New cdpDOMNode
            obj_children.init v
            children.Add obj_children
            Set obj_children = Nothing
        Next v
    End If
    If obj.Exists("attributes") Then
        For each v in obj.Item("attributes")
            attributes.Add v
        Next v
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
        contentDocument.init obj.Item("contentDocument")
    End If
    If obj.Exists("shadowRoots") Then
        For each v in obj.Item("shadowRoots")
            Dim obj_shadowRoots As cdpDOMNode: Set obj_shadowRoots = New cdpDOMNode
            obj_shadowRoots.init v
            shadowRoots.Add obj_shadowRoots
            Set obj_shadowRoots = Nothing
        Next v
    End If
    If obj.Exists("templateContent") Then
        Set templateContent = New cdpDOMNode
        templateContent.init obj.Item("templateContent")
    End If
    If obj.Exists("pseudoElements") Then
        For each v in obj.Item("pseudoElements")
            Dim obj_pseudoElements As cdpDOMNode: Set obj_pseudoElements = New cdpDOMNode
            obj_pseudoElements.init v
            pseudoElements.Add obj_pseudoElements
            Set obj_pseudoElements = Nothing
        Next v
    End If
    If obj.Exists("importedDocument") Then
        Set importedDocument = New cdpDOMNode
        importedDocument.init obj.Item("importedDocument")
    End If
    If obj.Exists("distributedNodes") Then
        For each v in obj.Item("distributedNodes")
            Dim obj_distributedNodes As cdpDOMBackendNode: Set obj_distributedNodes = New cdpDOMBackendNode
            obj_distributedNodes.init v
            distributedNodes.Add obj_distributedNodes
            Set obj_distributedNodes = Nothing
        Next v
    End If
    If obj.Exists("isSVG") Then Let isSVG = CBool(obj.Item("isSVG"))
    If obj.Exists("compatibilityMode") Then Let compatibilityMode = CStr(obj.Item("compatibilityMode"))

    If Not b Is Nothing Then 
        Set browser = b
        If nodeId > 0 Then Set resolveNode = browser.cdp.DOM.resolveNode(nodeId)
        If nodeId > 0 Then Set objectProperties = browser.cdp.Runtime.getProperties(resolveNode("objectId"))
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

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
    if documentURL <> vbNullString Then dict("documentURL") = documentURL
    if baseURL <> vbNullString Then dict("baseURL") = baseURL
    if publicId <> vbNullString Then dict("publicId") = publicId
    if systemId <> vbNullString Then dict("systemId") = systemId
    if internalSubset <> vbNullString Then dict("internalSubset") = internalSubset
    if xmlVersion <> vbNullString Then dict("xmlVersion") = xmlVersion
    if name <> vbNullString Then dict("name") = name
    if value <> vbNullString Then dict("value") = value
    if pseudoType <> vbNullString Then dict("pseudoType") = pseudoType
    if shadowRootType <> vbNullString Then dict("shadowRootType") = shadowRootType
    if frameId <> vbNullString Then dict("frameId") = frameId
    If Not contentDocument Is Nothing Then Set dict("contentDocument") = contentDocument.asDictionary()
    Set dict("shadowRoots") = shadowRoots
    If Not templateContent Is Nothing Then Set dict("templateContent") = templateContent.asDictionary()
    Set dict("pseudoElements") = pseudoElements
    If Not importedDocument Is Nothing Then Set dict("importedDocument") = importedDocument.asDictionary()
    Set dict("distributedNodes") = distributedNodes
    if Not IsEmpty(isSVG) Then dict("isSVG") = isSVG
    if compatibilityMode <> vbNullString Then dict("compatibilityMode") = compatibilityMode

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set children = New Collection
    Set attributes = New Collection
    Set shadowRoots = New Collection
    Set pseudoElements = New Collection
    Set distributedNodes = New Collection
End Sub

' ************* Convenience functions *********************
Public Sub click(Optional ByVal strategy As NavigationStrategy = NavigationStrategy.none)
    browser.click nodeId:=nodeId, backendNodeId:=backendNodeId, strategy:=strategy
End Sub

Public Sub getCenter(ByRef x As Double, ByRef y As Double)
    browser.getNodeCenter x, y, nodeId:=nodeId, backendNodeId:=backendNodeId
End Sub


Public Property Get elementValue() As String
    If value <> vbNullString Then
        elementValue = value
    ElseIf nodeValue <> vbNullString Then
        elementValue = nodeValue
    Else
        elementValue = elementAttribute("value")
    End If
End Property

Public Property Let elementValue(ByVal newvalue As String)
    setValue newvalue, SetNodeValueMode.SNV_AsAttribute
End Property

Public Function setValue(ByVal newvalue As String, Optional mode As SetNodeValueMode = SetNodeValueMode.SNV_AsAttribute)
    Select Case mode
        Case SetNodeValueMode.SNV_AsAttribute
            elementAttribute("value") = newvalue
        Case SetNodeValueMode.SNV_AsValue
            value = newvalue
        Case SetNodeValueMode.SNV_AsNodeValue
            nodeValue = newvalue
        Case SetNodeValueMode.SNV_AsFakeInput
            browser.cdp.SimulateInput.dispatchKeyEvent newvalue
    End Select
End Function


Public Property Get elementAttribute(ByVal name As String) As String
    elementAttribute = browser.getNodeAttribute(Me, name)
End Property

Public Property Let elementAttribute(ByVal name As String, ByVal value As String)
    browser.cdp.DOM.setAttributeValue nodeId, name, value
End Property


'nodeId >> resolveNode >> objectId >> objectProperties
'see Runtime.getProperties
Private Function getProperty(ByVal propertyName As String) As String
    Dim index As Integer
    Dim objectProperty As Dictionary
    
    For index = 1 To objectProperties("result").count
        Set objectProperty = objectProperties("result")(index)
        
        If objectProperty("name") = propertyName Then
            getProperty = objectProperty("value")("value")
            Exit For
        End If
    Next index
End Function

Public Function getIdOrName() As String
    getIdOrName = browser.getNodeIdOrName(Me)
End Function

Public Property Get id() As String
    id = getProperty("id")
End Property

Public Property Get className() As String
    className = getProperty("className")
End Property

Public Property Get title() As String
    title = getProperty("title")
End Property

Public Property Get textContent() As String
    textContent = getProperty("textContent")
End Property

Public Property Get tagName() As String
    tagName = getProperty("tagName")
End Property

Public Property Get outerText() As String
    outerText = getProperty("outerText")
End Property

Public Property Get innerText() As String
    innerText = getProperty("innerText")
End Property

Public Property Get innerHTML() As String
    innerHTML = getProperty("innerHTML")
End Property

Public Property Get outerHTML() As String
    outerHTML = getProperty("outerHTML")
End Property


'Parent
'NextSibling
'PrevSibling
