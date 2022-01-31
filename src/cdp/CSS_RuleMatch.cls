VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSRuleMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "RuleMatch"
' CSS.RuleMatch
' Match data for a CSS rule.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' rule : CSSRule
' CSS rule in the match.
Public rule AS cdpCSSCSSRule

' matchingSelectors : array of integer
' Matching selector indices in the rule's selectorList selectors (0-based).
Public matchingSelectors AS Collection

Public Function init(ByVal obj as Dictionary) As cdpCSSRuleMatch
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Set rule = New cdpCSSCSSRule
    rule.init obj.Item("rule")
    For each v in obj.Item("matchingSelectors")
        matchingSelectors.Add v
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("rule") = rule.asDictionary()
    Set dict("matchingSelectors") = matchingSelectors

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set matchingSelectors = New Collection
End Sub

