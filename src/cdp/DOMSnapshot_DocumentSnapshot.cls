VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDOMSnapshotDocumentSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DocumentSnapshot"
' DOMSnapshot.DocumentSnapshot
' Document snapshot.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' documentURL : StringIndex
' Document URL that `Document` or `FrameOwner` node points to.
Public documentURL AS Long

' title : StringIndex
' Document title.
Public title AS Long

' baseURL : StringIndex
' Base URL that `Document` or `FrameOwner` node uses for URL completion.
Public baseURL AS Long

' contentLanguage : StringIndex
' Contains the document's content language.
Public contentLanguage AS Long

' encodingName : StringIndex
' Contains the document's character set encoding.
Public encodingName AS Long

' publicId : StringIndex
' `DocumentType` node's publicId.
Public publicId AS Long

' systemId : StringIndex
' `DocumentType` node's systemId.
Public systemId AS Long

' frameId : StringIndex
' Frame ID for frame owner elements and also for the document node.
Public frameId AS Long

' nodes : NodeTreeSnapshot
' A table with dom nodes.
Public nodes AS cdpDOMSnapshotNodeTreeSnapshot

' layout : LayoutTreeSnapshot
' The nodes in the layout tree.
Public layout AS cdpDOMSnapshotLayoutTreeSnapsho

' textBoxes : TextBoxSnapshot
' The post-layout inline text nodes.
Public textBoxes AS cdpDOMSnapshotTextBoxSnapshot

' scrollOffsetX : number
' Optional
' Horizontal scroll offset.
Public scrollOffsetX AS Double

' scrollOffsetY : number
' Optional
' Vertical scroll offset.
Public scrollOffsetY AS Double

' contentWidth : number
' Optional
' Document content width.
Public contentWidth AS Double

' contentHeight : number
' Optional
' Document content height.
Public contentHeight AS Double

Public Function init(ByVal obj as Dictionary) As cdpDOMSnapshotDocumentSnapshot
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let documentURL = CLng(obj.Item("documentURL"))
    Let title = CLng(obj.Item("title"))
    Let baseURL = CLng(obj.Item("baseURL"))
    Let contentLanguage = CLng(obj.Item("contentLanguage"))
    Let encodingName = CLng(obj.Item("encodingName"))
    Let publicId = CLng(obj.Item("publicId"))
    Let systemId = CLng(obj.Item("systemId"))
    Let frameId = CLng(obj.Item("frameId"))
    Set nodes = New cdpDOMSnapshotNodeTreeSnapshot
    nodes.init obj.Item("nodes")
    Set layout = New cdpDOMSnapshotLayoutTreeSnapsho
    layout.init obj.Item("layout")
    Set textBoxes = New cdpDOMSnapshotTextBoxSnapshot
    textBoxes.init obj.Item("textBoxes")
    If obj.Exists("scrollOffsetX") Then Let scrollOffsetX = CDbl(obj.Item("scrollOffsetX"))
    If obj.Exists("scrollOffsetY") Then Let scrollOffsetY = CDbl(obj.Item("scrollOffsetY"))
    If obj.Exists("contentWidth") Then Let contentWidth = CDbl(obj.Item("contentWidth"))
    If obj.Exists("contentHeight") Then Let contentHeight = CDbl(obj.Item("contentHeight"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("documentURL") = documentURL
    dict("title") = title
    dict("baseURL") = baseURL
    dict("contentLanguage") = contentLanguage
    dict("encodingName") = encodingName
    dict("publicId") = publicId
    dict("systemId") = systemId
    dict("frameId") = frameId
    Set dict("nodes") = nodes.asDictionary()
    Set dict("layout") = layout.asDictionary()
    Set dict("textBoxes") = textBoxes.asDictionary()
    dict("scrollOffsetX") = scrollOffsetX
    dict("scrollOffsetY") = scrollOffsetY
    dict("contentWidth") = contentWidth
    dict("contentHeight") = contentHeight

    Set asDictionary = dict
End Function

