VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotDOMNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DOMNode"
' DOMSnapshot.DOMNode
' A Node in the DOM tree.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' nodeType : integer
' `Node`'s nodeType.
Public nodeType AS Long

' nodeName : string
' `Node`'s nodeName.
Public nodeName AS string

' nodeValue : string
' `Node`'s nodeValue.
Public nodeValue AS string

' textValue : string
' Optional
' Only set for textarea elements, contains the text value.
Public textValue AS string

' inputValue : string
' Optional
' Only set for input elements, contains the input's associated text value.
Public inputValue AS string

' inputChecked : boolean
' Optional
' Only set for radio and checkbox input elements, indicates if the element has been checked
Public inputChecked AS Variant ' boolean

' optionSelected : boolean
' Optional
' Only set for option elements, indicates if the element has been selected
Public optionSelected AS Variant ' boolean

' backendNodeId : DOM.BackendNodeId
' `Node`'s id, corresponds to DOM.Node.backendNodeId.
Public backendNodeId AS Long

' childNodeIndexes : array of integer
' Optional
' The indexes of the node's child nodes in the `domNodes` array returned by `getSnapshot`, if
    '   any.
Public childNodeIndexes AS Collection

' attributes : array of cdpDOMSnapshotNameValue
' Optional
' Attributes of an `Element` node.
Public attributes AS Collection

' pseudoElementIndexes : array of integer
' Optional
' Indexes of pseudo elements associated with this node in the `domNodes` array returned by
    '   `getSnapshot`, if any.
Public pseudoElementIndexes AS Collection

' layoutNodeIndex : integer
' Optional
' The index of the node's related layout tree node in the `layoutTreeNodes` array returned by
    '   `getSnapshot`, if any.
Public layoutNodeIndex AS Long

' documentURL : string
' Optional
' Document URL that `Document` or `FrameOwner` node points to.
Public documentURL AS string

' baseURL : string
' Optional
' Base URL that `Document` or `FrameOwner` node uses for URL completion.
Public baseURL AS string

' contentLanguage : string
' Optional
' Only set for documents, contains the document's content language.
Public contentLanguage AS string

' documentEncoding : string
' Optional
' Only set for documents, contains the document's character set encoding.
Public documentEncoding AS string

' publicId : string
' Optional
' `DocumentType` node's publicId.
Public publicId AS string

' systemId : string
' Optional
' `DocumentType` node's systemId.
Public systemId AS string

' frameId : Page.FrameId
' Optional
' Frame ID for frame owner elements and also for the document node.
Public frameId AS string

' contentDocumentIndex : integer
' Optional
' The index of a frame owner element's content document in the `domNodes` array returned by
    '   `getSnapshot`, if any.
Public contentDocumentIndex AS Long

' pseudoType : DOM.PseudoType
' Optional
' Type of a pseudo element node.
Public pseudoType AS string

' shadowRootType : DOM.ShadowRootType
' Optional
' Shadow root type.
Public shadowRootType AS string

' isClickable : boolean
' Optional
' Whether this DOM node responds to mouse clicks. This includes nodes that have had click
    '   event listeners attached via JavaScript as well as anchor tags that naturally navigate when
    '   clicked.
Public isClickable AS Variant ' boolean

' eventListeners : array of cdpDOMDebuggerEventListener
' Optional
' Details of the node's event listeners, if any.
Public eventListeners AS Collection

' currentSourceURL : string
' Optional
' The selected url for nodes with a srcset attribute.
Public currentSourceURL AS string

' originURL : string
' Optional
' The url of the script (if any) that generates this node.
Public originURL AS string

' scrollOffsetX : number
' Optional
' Scroll offsets, set when this node is a Document.
Public scrollOffsetX AS Double

' scrollOffsetY : number
' Optional
' 
Public scrollOffsetY AS Double

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotDOMNode
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let nodeType = CLng(obj.Item("nodeType"))
    Let nodeName = CStr(obj.Item("nodeName"))
    Let nodeValue = CStr(obj.Item("nodeValue"))
    If obj.Exists("textValue") Then Let textValue = CStr(obj.Item("textValue"))
    If obj.Exists("inputValue") Then Let inputValue = CStr(obj.Item("inputValue"))
    If obj.Exists("inputChecked") Then Let inputChecked = CBool(obj.Item("inputChecked"))
    If obj.Exists("optionSelected") Then Let optionSelected = CBool(obj.Item("optionSelected"))
    Let backendNodeId = CLng(obj.Item("backendNodeId"))
    If obj.Exists("childNodeIndexes") Then
        For each v in obj.Item("childNodeIndexes")
            childNodeIndexes.Add v
        Next v
    End If
    If obj.Exists("attributes") Then
        For each v in obj.Item("attributes")
            Dim obj_attributes As cdpDOMSnapshotNameValue: Set obj_attributes = New cdpDOMSnapshotNameValue
            obj_attributes.init v
            attributes.Add obj_attributes
            Set obj_attributes = Nothing
        Next v
    End If
    If obj.Exists("pseudoElementIndexes") Then
        For each v in obj.Item("pseudoElementIndexes")
            pseudoElementIndexes.Add v
        Next v
    End If
    If obj.Exists("layoutNodeIndex") Then Let layoutNodeIndex = CLng(obj.Item("layoutNodeIndex"))
    If obj.Exists("documentURL") Then Let documentURL = CStr(obj.Item("documentURL"))
    If obj.Exists("baseURL") Then Let baseURL = CStr(obj.Item("baseURL"))
    If obj.Exists("contentLanguage") Then Let contentLanguage = CStr(obj.Item("contentLanguage"))
    If obj.Exists("documentEncoding") Then Let documentEncoding = CStr(obj.Item("documentEncoding"))
    If obj.Exists("publicId") Then Let publicId = CStr(obj.Item("publicId"))
    If obj.Exists("systemId") Then Let systemId = CStr(obj.Item("systemId"))
    If obj.Exists("frameId") Then Let frameId = CStr(obj.Item("frameId"))
    If obj.Exists("contentDocumentIndex") Then Let contentDocumentIndex = CLng(obj.Item("contentDocumentIndex"))
    If obj.Exists("pseudoType") Then Let pseudoType = CStr(obj.Item("pseudoType"))
    If obj.Exists("shadowRootType") Then Let shadowRootType = CStr(obj.Item("shadowRootType"))
    If obj.Exists("isClickable") Then Let isClickable = CBool(obj.Item("isClickable"))
    If obj.Exists("eventListeners") Then
        For each v in obj.Item("eventListeners")
            Dim obj_eventListeners As cdpDOMDebuggerEventListener: Set obj_eventListeners = New cdpDOMDebuggerEventListener
            obj_eventListeners.init v
            eventListeners.Add obj_eventListeners
            Set obj_eventListeners = Nothing
        Next v
    End If
    If obj.Exists("currentSourceURL") Then Let currentSourceURL = CStr(obj.Item("currentSourceURL"))
    If obj.Exists("originURL") Then Let originURL = CStr(obj.Item("originURL"))
    If obj.Exists("scrollOffsetX") Then Let scrollOffsetX = CDbl(obj.Item("scrollOffsetX"))
    If obj.Exists("scrollOffsetY") Then Let scrollOffsetY = CDbl(obj.Item("scrollOffsetY"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("nodeType") = nodeType
    dict("nodeName") = nodeName
    dict("nodeValue") = nodeValue
    if textValue <> vbNullString Then dict("textValue") = textValue
    if inputValue <> vbNullString Then dict("inputValue") = inputValue
    if Not IsEmpty(inputChecked) Then dict("inputChecked") = inputChecked
    if Not IsEmpty(optionSelected) Then dict("optionSelected") = optionSelected
    dict("backendNodeId") = backendNodeId
    Set dict("childNodeIndexes") = childNodeIndexes
    Set dict("attributes") = attributes
    Set dict("pseudoElementIndexes") = pseudoElementIndexes
    dict("layoutNodeIndex") = layoutNodeIndex
    if documentURL <> vbNullString Then dict("documentURL") = documentURL
    if baseURL <> vbNullString Then dict("baseURL") = baseURL
    if contentLanguage <> vbNullString Then dict("contentLanguage") = contentLanguage
    if documentEncoding <> vbNullString Then dict("documentEncoding") = documentEncoding
    if publicId <> vbNullString Then dict("publicId") = publicId
    if systemId <> vbNullString Then dict("systemId") = systemId
    if frameId <> vbNullString Then dict("frameId") = frameId
    dict("contentDocumentIndex") = contentDocumentIndex
    if pseudoType <> vbNullString Then dict("pseudoType") = pseudoType
    if shadowRootType <> vbNullString Then dict("shadowRootType") = shadowRootType
    if Not IsEmpty(isClickable) Then dict("isClickable") = isClickable
    Set dict("eventListeners") = eventListeners
    if currentSourceURL <> vbNullString Then dict("currentSourceURL") = currentSourceURL
    if originURL <> vbNullString Then dict("originURL") = originURL
    dict("scrollOffsetX") = scrollOffsetX
    dict("scrollOffsetY") = scrollOffsetY

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set childNodeIndexes = New Collection
    Set attributes = New Collection
    Set pseudoElementIndexes = New Collection
    Set eventListeners = New Collection
End Sub

