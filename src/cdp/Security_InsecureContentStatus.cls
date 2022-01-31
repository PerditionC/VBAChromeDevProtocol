VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSecurityInsecureContentStatu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "InsecureContentStatus"
' Security.InsecureContentStatus
' Information about insecure content on the page.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' ranMixedContent : boolean
' Always false.
Public ranMixedContent AS boolean

' displayedMixedContent : boolean
' Always false.
Public displayedMixedContent AS boolean

' containedMixedForm : boolean
' Always false.
Public containedMixedForm AS boolean

' ranContentWithCertErrors : boolean
' Always false.
Public ranContentWithCertErrors AS boolean

' displayedContentWithCertErrors : boolean
' Always false.
Public displayedContentWithCertErrors AS boolean

' ranInsecureContentStyle : SecurityState
' Always set to unknown.
Public ranInsecureContentStyle AS string

' displayedInsecureContentStyle : SecurityState
' Always set to unknown.
Public displayedInsecureContentStyle AS string

Public Function init(ByVal obj as Dictionary) As cdpSecurityInsecureContentStatu
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let ranMixedContent = CBool(obj.Item("ranMixedContent"))
    Let displayedMixedContent = CBool(obj.Item("displayedMixedContent"))
    Let containedMixedForm = CBool(obj.Item("containedMixedForm"))
    Let ranContentWithCertErrors = CBool(obj.Item("ranContentWithCertErrors"))
    Let displayedContentWithCertErrors = CBool(obj.Item("displayedContentWithCertErrors"))
    Let ranInsecureContentStyle = CStr(obj.Item("ranInsecureContentStyle"))
    Let displayedInsecureContentStyle = CStr(obj.Item("displayedInsecureContentStyle"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("ranMixedContent") = ranMixedContent
    dict("displayedMixedContent") = displayedMixedContent
    dict("containedMixedForm") = containedMixedForm
    dict("ranContentWithCertErrors") = ranContentWithCertErrors
    dict("displayedContentWithCertErrors") = displayedContentWithCertErrors
    dict("ranInsecureContentStyle") = ranInsecureContentStyle
    dict("displayedInsecureContentStyle") = displayedInsecureContentStyle

    Set asDictionary = dict
End Function

