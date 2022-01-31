VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSInheritedStyleEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "InheritedStyleEntry"
' CSS.InheritedStyleEntry
' Inherited CSS rule collection from ancestor node.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' inlineStyle : CSSStyle
' Optional
' The ancestor node's inline style, if any, in the style inheritance chain.
Public inlineStyle AS cdpCSSCSSStyle

' matchedCSSRules : array of cdpCSSRuleMatch
' Matches of CSS rules matching the ancestor node in the style inheritance chain.
Public matchedCSSRules AS Collection

Public Function init(ByVal obj as Dictionary) As cdpCSSInheritedStyleEntry
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    If obj.Exists("inlineStyle") Then
        Set inlineStyle = New cdpCSSCSSStyle
        inlineStyle.init obj.Item("inlineStyle")
    End If
    For each v in obj.Item("matchedCSSRules")
        Dim obj_matchedCSSRules As cdpCSSRuleMatch: Set obj_matchedCSSRules = New cdpCSSRuleMatch
        obj_matchedCSSRules.init v
        matchedCSSRules.Add obj_matchedCSSRules
        Set obj_matchedCSSRules = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    If Not inlineStyle Is Nothing Then Set dict("inlineStyle") = inlineStyle.asDictionary()
    Set dict("matchedCSSRules") = matchedCSSRules

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set matchedCSSRules = New Collection
End Sub

