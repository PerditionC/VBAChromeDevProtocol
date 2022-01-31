VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotNodeTreeSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "NodeTreeSnapshot"
' DOMSnapshot.NodeTreeSnapshot
' Table containing nodes.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' parentIndex : array of integer
' Optional
' Parent node index.
Public parentIndex AS Collection

' nodeType : array of integer
' Optional
' `Node`'s nodeType.
Public nodeType AS Collection

' shadowRootType : RareStringData
' Optional
' Type of the shadow root the `Node` is in. String values are equal to the `ShadowRootType` enum.
Public shadowRootType AS cdpDOMSnapshotRareStringData

' nodeName : array of cdpDOMSnapshotStringIndex
' Optional
' `Node`'s nodeName.
Public nodeName AS Collection

' nodeValue : array of cdpDOMSnapshotStringIndex
' Optional
' `Node`'s nodeValue.
Public nodeValue AS Collection

' backendNodeId : array of cdpDOMBackendNodeId
' Optional
' `Node`'s id, corresponds to DOM.Node.backendNodeId.
Public backendNodeId AS Collection

' attributes : array of cdpDOMSnapshotArrayOfStrings
' Optional
' Attributes of an `Element` node. Flatten name, value pairs.
Public attributes AS Collection

' textValue : RareStringData
' Optional
' Only set for textarea elements, contains the text value.
Public textValue AS cdpDOMSnapshotRareStringData

' inputValue : RareStringData
' Optional
' Only set for input elements, contains the input's associated text value.
Public inputValue AS cdpDOMSnapshotRareStringData

' inputChecked : RareBooleanData
' Optional
' Only set for radio and checkbox input elements, indicates if the element has been checked
Public inputChecked AS cdpDOMSnapshotRareBooleanData

' optionSelected : RareBooleanData
' Optional
' Only set for option elements, indicates if the element has been selected
Public optionSelected AS cdpDOMSnapshotRareBooleanData

' contentDocumentIndex : RareIntegerData
' Optional
' The index of the document in the list of the snapshot documents.
Public contentDocumentIndex AS cdpDOMSnapshotRareIntegerData

' pseudoType : RareStringData
' Optional
' Type of a pseudo element node.
Public pseudoType AS cdpDOMSnapshotRareStringData

' isClickable : RareBooleanData
' Optional
' Whether this DOM node responds to mouse clicks. This includes nodes that have had click
    '   event listeners attached via JavaScript as well as anchor tags that naturally navigate when
    '   clicked.
Public isClickable AS cdpDOMSnapshotRareBooleanData

' currentSourceURL : RareStringData
' Optional
' The selected url for nodes with a srcset attribute.
Public currentSourceURL AS cdpDOMSnapshotRareStringData

' originURL : RareStringData
' Optional
' The url of the script (if any) that generates this node.
Public originURL AS cdpDOMSnapshotRareStringData

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotNodeTreeSnapshot
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    If obj.Exists("parentIndex") Then
        For each v in obj.Item("parentIndex")
            parentIndex.Add v
        Next v
    End If
    If obj.Exists("nodeType") Then
        For each v in obj.Item("nodeType")
            nodeType.Add v
        Next v
    End If
    If obj.Exists("shadowRootType") Then
        Set shadowRootType = New cdpDOMSnapshotRareStringData
        shadowRootType.init obj.Item("shadowRootType")
    End If
    If obj.Exists("nodeName") Then
        For each v in obj.Item("nodeName")
            nodeName.Add v
        Next v
    End If
    If obj.Exists("nodeValue") Then
        For each v in obj.Item("nodeValue")
            nodeValue.Add v
        Next v
    End If
    If obj.Exists("backendNodeId") Then
        For each v in obj.Item("backendNodeId")
            backendNodeId.Add v
        Next v
    End If
    If obj.Exists("attributes") Then
        For each v in obj.Item("attributes")
            attributes.Add v
        Next v
    End If
    If obj.Exists("textValue") Then
        Set textValue = New cdpDOMSnapshotRareStringData
        textValue.init obj.Item("textValue")
    End If
    If obj.Exists("inputValue") Then
        Set inputValue = New cdpDOMSnapshotRareStringData
        inputValue.init obj.Item("inputValue")
    End If
    If obj.Exists("inputChecked") Then
        Set inputChecked = New cdpDOMSnapshotRareBooleanData
        inputChecked.init obj.Item("inputChecked")
    End If
    If obj.Exists("optionSelected") Then
        Set optionSelected = New cdpDOMSnapshotRareBooleanData
        optionSelected.init obj.Item("optionSelected")
    End If
    If obj.Exists("contentDocumentIndex") Then
        Set contentDocumentIndex = New cdpDOMSnapshotRareIntegerData
        contentDocumentIndex.init obj.Item("contentDocumentIndex")
    End If
    If obj.Exists("pseudoType") Then
        Set pseudoType = New cdpDOMSnapshotRareStringData
        pseudoType.init obj.Item("pseudoType")
    End If
    If obj.Exists("isClickable") Then
        Set isClickable = New cdpDOMSnapshotRareBooleanData
        isClickable.init obj.Item("isClickable")
    End If
    If obj.Exists("currentSourceURL") Then
        Set currentSourceURL = New cdpDOMSnapshotRareStringData
        currentSourceURL.init obj.Item("currentSourceURL")
    End If
    If obj.Exists("originURL") Then
        Set originURL = New cdpDOMSnapshotRareStringData
        originURL.init obj.Item("originURL")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("parentIndex") = parentIndex
    Set dict("nodeType") = nodeType
    If Not shadowRootType Is Nothing Then Set dict("shadowRootType") = shadowRootType.asDictionary()
    Set dict("nodeName") = nodeName
    Set dict("nodeValue") = nodeValue
    Set dict("backendNodeId") = backendNodeId
    Set dict("attributes") = attributes
    If Not textValue Is Nothing Then Set dict("textValue") = textValue.asDictionary()
    If Not inputValue Is Nothing Then Set dict("inputValue") = inputValue.asDictionary()
    If Not inputChecked Is Nothing Then Set dict("inputChecked") = inputChecked.asDictionary()
    If Not optionSelected Is Nothing Then Set dict("optionSelected") = optionSelected.asDictionary()
    If Not contentDocumentIndex Is Nothing Then Set dict("contentDocumentIndex") = contentDocumentIndex.asDictionary()
    If Not pseudoType Is Nothing Then Set dict("pseudoType") = pseudoType.asDictionary()
    If Not isClickable Is Nothing Then Set dict("isClickable") = isClickable.asDictionary()
    If Not currentSourceURL Is Nothing Then Set dict("currentSourceURL") = currentSourceURL.asDictionary()
    If Not originURL Is Nothing Then Set dict("originURL") = originURL.asDictionary()

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set parentIndex = New Collection
    Set nodeType = New Collection
    Set nodeName = New Collection
    Set nodeValue = New Collection
    Set backendNodeId = New Collection
    Set attributes = New Collection
End Sub

