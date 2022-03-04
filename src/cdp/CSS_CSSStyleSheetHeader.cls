VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSCSSStyleSheetHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CSSStyleSheetHeader"
' CSS.CSSStyleSheetHeader
' CSS stylesheet metainformation.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' styleSheetId : StyleSheetId
' The stylesheet identifier.
Public styleSheetId AS string

' frameId : Page.FrameId
' Owner frame identifier.
Public frameId AS string

' sourceURL : string
' Stylesheet resource URL. Empty if this is a constructed stylesheet created using
    '   new CSSStyleSheet() (but non-empty if this is a constructed sylesheet imported
    '   as a CSS module script).
Public sourceURL AS string

' sourceMapURL : string
' Optional
' URL of source map associated with the stylesheet (if any).
Public sourceMapURL AS string

' origin : StyleSheetOrigin
' Stylesheet origin.
Public origin AS string

' title : string
' Stylesheet title.
Public title AS string

' ownerNode : DOM.BackendNodeId
' Optional
' The backend id for the owner node of the stylesheet.
Public ownerNode AS Long

' disabled : boolean
' Denotes whether the stylesheet is disabled.
Public disabled AS boolean

' hasSourceURL : boolean
' Optional
' Whether the sourceURL field value comes from the sourceURL comment.
Public hasSourceURL AS Variant ' boolean

' isInline : boolean
' Whether this stylesheet is created for STYLE tag by parser. This flag is not set for
    '   document.written STYLE tags.
Public isInline AS boolean

' isMutable : boolean
' Whether this stylesheet is mutable. Inline stylesheets become mutable
    '   after they have been modified via CSSOM API.
    '   <link> element's stylesheets become mutable only if DevTools modifies them.
    '   Constructed stylesheets (new CSSStyleSheet()) are mutable immediately after creation.
Public isMutable AS boolean

' isConstructed : boolean
' True if this stylesheet is created through new CSSStyleSheet() or imported as a
    '   CSS module script.
Public isConstructed AS boolean

' startLine : number
' Line offset of the stylesheet within the resource (zero based).
Public startLine AS Double

' startColumn : number
' Column offset of the stylesheet within the resource (zero based).
Public startColumn AS Double

' length : number
' Size of the content (in characters).
Public length AS Double

' endLine : number
' Line offset of the end of the stylesheet within the resource (zero based).
Public endLine AS Double

' endColumn : number
' Column offset of the end of the stylesheet within the resource (zero based).
Public endColumn AS Double

Public Function init(ByVal obj as Dictionary) As cdpCSSCSSStyleSheetHeader
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let styleSheetId = CStr(obj.Item("styleSheetId"))
    Let frameId = CStr(obj.Item("frameId"))
    Let sourceURL = CStr(obj.Item("sourceURL"))
    If obj.Exists("sourceMapURL") Then Let sourceMapURL = CStr(obj.Item("sourceMapURL"))
    Let origin = CStr(obj.Item("origin"))
    Let title = CStr(obj.Item("title"))
    If obj.Exists("ownerNode") Then Let ownerNode = CLng(obj.Item("ownerNode"))
    Let disabled = CBool(obj.Item("disabled"))
    If obj.Exists("hasSourceURL") Then Let hasSourceURL = CBool(obj.Item("hasSourceURL"))
    Let isInline = CBool(obj.Item("isInline"))
    Let isMutable = CBool(obj.Item("isMutable"))
    Let isConstructed = CBool(obj.Item("isConstructed"))
    Let startLine = CDbl(obj.Item("startLine"))
    Let startColumn = CDbl(obj.Item("startColumn"))
    Let length = CDbl(obj.Item("length"))
    Let endLine = CDbl(obj.Item("endLine"))
    Let endColumn = CDbl(obj.Item("endColumn"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("styleSheetId") = styleSheetId
    dict("frameId") = frameId
    dict("sourceURL") = sourceURL
    if sourceMapURL <> vbNullString Then dict("sourceMapURL") = sourceMapURL
    dict("origin") = origin
    dict("title") = title
    dict("ownerNode") = ownerNode
    dict("disabled") = disabled
    if Not IsEmpty(hasSourceURL) Then dict("hasSourceURL") = hasSourceURL
    dict("isInline") = isInline
    dict("isMutable") = isMutable
    dict("isConstructed") = isConstructed
    dict("startLine") = startLine
    dict("startColumn") = startColumn
    dict("length") = length
    dict("endLine") = endLine
    dict("endColumn") = endColumn

    Set asDictionary = dict
End Function

