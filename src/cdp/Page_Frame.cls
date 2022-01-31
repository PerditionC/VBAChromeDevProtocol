VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Frame"
' Page.Frame
' Information about the Frame on the page.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' id : FrameId
' Frame unique identifier.
Public id AS string

' parentId : FrameId
' Optional
' Parent frame identifier.
Public parentId AS string

' loaderId : Network.LoaderId
' Identifier of the loader associated with this frame.
Public loaderId AS string

' name : string
' Optional
' Frame's name as specified in the tag.
Public name AS string

' url : string
' Frame document's URL without fragment.
Public url AS string

' urlFragment : string
' Optional
' Frame document's URL fragment including the '#'.
Public urlFragment AS string

' domainAndRegistry : string
' Frame document's registered domain, taking the public suffixes list into account.
    '   Extracted from the Frame's url.
    '   Example URLs: http://www.google.com/file.html -> "google.com"
    '                 http://a.b.co.uk/file.html      -> "b.co.uk"
Public domainAndRegistry AS string

' securityOrigin : string
' Frame document's security origin.
Public securityOrigin AS string

' mimeType : string
' Frame document's mimeType as determined by the browser.
Public mimeType AS string

' unreachableUrl : string
' Optional
' If the frame failed to load, this contains the URL that could not be loaded. Note that unlike url above, this URL may contain a fragment.
Public unreachableUrl AS string

' adFrameStatus : AdFrameStatus
' Optional
' Indicates whether this frame was tagged as an ad and why.
Public adFrameStatus AS cdpPageAdFrameStatus

' secureContextType : SecureContextType
' Indicates whether the main document is a secure context and explains why that is the case.
Public secureContextType AS string

' crossOriginIsolatedContextType : CrossOriginIsolatedContextType
' Indicates whether this is a cross origin isolated context.
Public crossOriginIsolatedContextType AS string

' gatedAPIFeatures : array of String
' Indicated which gated APIs / features are available.
Public gatedAPIFeatures AS Collection

Public Function init(ByVal obj as Dictionary) As cdpPageFrame
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let id = CStr(obj.Item("id"))
    If obj.Exists("parentId") Then Let parentId = CStr(obj.Item("parentId"))
    Let loaderId = CStr(obj.Item("loaderId"))
    If obj.Exists("name") Then Let name = CStr(obj.Item("name"))
    Let url = CStr(obj.Item("url"))
    If obj.Exists("urlFragment") Then Let urlFragment = CStr(obj.Item("urlFragment"))
    Let domainAndRegistry = CStr(obj.Item("domainAndRegistry"))
    Let securityOrigin = CStr(obj.Item("securityOrigin"))
    Let mimeType = CStr(obj.Item("mimeType"))
    If obj.Exists("unreachableUrl") Then Let unreachableUrl = CStr(obj.Item("unreachableUrl"))
    If obj.Exists("adFrameStatus") Then
        Set adFrameStatus = New cdpPageAdFrameStatus
        adFrameStatus.init obj.Item("adFrameStatus")
    End If
    Let secureContextType = CStr(obj.Item("secureContextType"))
    Let crossOriginIsolatedContextType = CStr(obj.Item("crossOriginIsolatedContextType"))
    For each v in obj.Item("gatedAPIFeatures")
        gatedAPIFeatures.Add v
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("id") = id
    if parentId <> vbNullString Then dict("parentId") = parentId
    dict("loaderId") = loaderId
    if name <> vbNullString Then dict("name") = name
    dict("url") = url
    if urlFragment <> vbNullString Then dict("urlFragment") = urlFragment
    dict("domainAndRegistry") = domainAndRegistry
    dict("securityOrigin") = securityOrigin
    dict("mimeType") = mimeType
    if unreachableUrl <> vbNullString Then dict("unreachableUrl") = unreachableUrl
    If Not adFrameStatus Is Nothing Then Set dict("adFrameStatus") = adFrameStatus.asDictionary()
    dict("secureContextType") = secureContextType
    dict("crossOriginIsolatedContextType") = crossOriginIsolatedContextType
    Set dict("gatedAPIFeatures") = gatedAPIFeatures

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set gatedAPIFeatures = New Collection
End Sub

