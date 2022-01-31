VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSMediaQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "MediaQuery"
' CSS.MediaQuery
' Media query descriptor.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' expressions : array of cdpCSSMediaQueryExpression
' Array of media query expressions.
Public expressions AS Collection

' active : boolean
' Whether the media query condition is satisfied.
Public active AS boolean

Public Function init(ByVal obj as Dictionary) As cdpCSSMediaQuery
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("expressions")
        Dim obj_expressions As cdpCSSMediaQueryExpression: Set obj_expressions = New cdpCSSMediaQueryExpression
        obj_expressions.init v
        expressions.Add obj_expressions
        Set obj_expressions = Nothing
    Next v
    Let active = CBool(obj.Item("active"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("expressions") = expressions
    dict("active") = active

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set expressions = New Collection
End Sub

