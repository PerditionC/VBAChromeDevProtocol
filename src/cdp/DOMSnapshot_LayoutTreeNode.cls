VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotLayoutTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "LayoutTreeNode"
' DOMSnapshot.LayoutTreeNode
' Details of an element in the DOM tree with a LayoutObject.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' domNodeIndex : integer
' The index of the related DOM node in the `domNodes` array returned by `getSnapshot`.
Public domNodeIndex AS Long

' boundingBox : DOM.Rect
' The bounding box in document coordinates. Note that scroll offset of the document is ignored.
Public boundingBox AS cdpDOMRect

' layoutText : string
' Optional
' Contents of the LayoutText, if any.
Public layoutText AS string

' inlineTextNodes : array of cdpDOMSnapshotInlineTextBox
' Optional
' The post-layout inline text nodes, if any.
Public inlineTextNodes AS Collection

' styleIndex : integer
' Optional
' Index into the `computedStyles` array returned by `getSnapshot`.
Public styleIndex AS Long

' paintOrder : integer
' Optional
' Global paint order index, which is determined by the stacking order of the nodes. Nodes
    '   that are painted together will have the same index. Only provided if includePaintOrder in
    '   getSnapshot was true.
Public paintOrder AS Long

' isStackingContext : boolean
' Optional
' Set to true to indicate the element begins a new stacking context.
Public isStackingContext AS Variant ' boolean

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotLayoutTreeNode
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let domNodeIndex = CLng(obj.Item("domNodeIndex"))
    Set boundingBox = New cdpDOMRect
    boundingBox.init obj.Item("boundingBox")
    If obj.Exists("layoutText") Then Let layoutText = CStr(obj.Item("layoutText"))
    If obj.Exists("inlineTextNodes") Then
        For each v in obj.Item("inlineTextNodes")
            Dim obj_inlineTextNodes As cdpDOMSnapshotInlineTextBox: Set obj_inlineTextNodes = New cdpDOMSnapshotInlineTextBox
            obj_inlineTextNodes.init v
            inlineTextNodes.Add obj_inlineTextNodes
            Set obj_inlineTextNodes = Nothing
        Next v
    End If
    If obj.Exists("styleIndex") Then Let styleIndex = CLng(obj.Item("styleIndex"))
    If obj.Exists("paintOrder") Then Let paintOrder = CLng(obj.Item("paintOrder"))
    If obj.Exists("isStackingContext") Then Let isStackingContext = CBool(obj.Item("isStackingContext"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("domNodeIndex") = domNodeIndex
    Set dict("boundingBox") = boundingBox.asDictionary()
    if layoutText <> vbNullString Then dict("layoutText") = layoutText
    Set dict("inlineTextNodes") = inlineTextNodes
    dict("styleIndex") = styleIndex
    dict("paintOrder") = paintOrder
    if Not IsEmpty(isStackingContext) Then dict("isStackingContext") = isStackingContext

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set inlineTextNodes = New Collection
End Sub

