VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSRuleUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "RuleUsage"
' CSS.RuleUsage
' CSS coverage information.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' styleSheetId : StyleSheetId
' The css style sheet identifier (absent for user agent stylesheet and user-specified
    '   stylesheet rules) this rule came from.
Public styleSheetId AS string

' startOffset : number
' Offset of the start of the rule (including selector) from the beginning of the stylesheet.
Public startOffset AS Double

' endOffset : number
' Offset of the end of the rule body from the beginning of the stylesheet.
Public endOffset AS Double

' used : boolean
' Indicates whether the rule was actually used by some element in the page.
Public used AS boolean

Public Function init(ByVal obj as Dictionary) As cdpCSSRuleUsage
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let styleSheetId = CStr(obj.Item("styleSheetId"))
    Let startOffset = CDbl(obj.Item("startOffset"))
    Let endOffset = CDbl(obj.Item("endOffset"))
    Let used = CBool(obj.Item("used"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("styleSheetId") = styleSheetId
    dict("startOffset") = startOffset
    dict("endOffset") = endOffset
    dict("used") = used

    Set asDictionary = dict
End Function

