VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSStyleDeclarationEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "StyleDeclarationEdit"
' CSS.StyleDeclarationEdit
' A descriptor of operation to mutate style declaration text.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' styleSheetId : StyleSheetId
' The css style sheet identifier.
Public styleSheetId AS string

' range : SourceRange
' The range of the style text in the enclosing stylesheet.
Public range AS cdpCSSSourceRange

' text : string
' New style text.
Public text AS string

Public Function init(ByVal obj as Dictionary) As cdpCSSStyleDeclarationEdit
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let styleSheetId = CStr(obj.Item("styleSheetId"))
    Set range = New cdpCSSSourceRange
    range.init obj.Item("range")
    Let text = CStr(obj.Item("text"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("styleSheetId") = styleSheetId
    Set dict("range") = range.asDictionary()
    dict("text") = text

    Set asDictionary = dict
End Function

