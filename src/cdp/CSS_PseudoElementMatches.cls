VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSPseudoElementMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PseudoElementMatches"
' CSS.PseudoElementMatches
' CSS rule collection for a single pseudo style.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' pseudoType : DOM.PseudoType
' Pseudo element type.
Public pseudoType AS string

' matches : array of cdpCSSRuleMatch
' Matches of CSS rules applicable to the pseudo style.
Public matches AS Collection

Public Function init(ByVal obj as Dictionary) As cdpCSSPseudoElementMatches
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let pseudoType = CStr(obj.Item("pseudoType"))
    For each v in obj.Item("matches")
        Dim obj_matches As cdpCSSRuleMatch: Set obj_matches = New cdpCSSRuleMatch
        obj_matches.init v
        matches.Add obj_matches
        Set obj_matches = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("pseudoType") = pseudoType
    Set dict("matches") = matches

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set matches = New Collection
End Sub

