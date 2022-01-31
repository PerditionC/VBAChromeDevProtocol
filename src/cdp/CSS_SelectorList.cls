VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSSelectorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SelectorList"
' CSS.SelectorList
' Selector list data.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' selectors : array of cdpCSSValue
' Selectors in the list.
Public selectors AS Collection

' text : string
' Rule selector text.
Public text AS string

Public Function init(ByVal obj as Dictionary) As cdpCSSSelectorList
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("selectors")
        Dim obj_selectors As cdpCSSValue: Set obj_selectors = New cdpCSSValue
        obj_selectors.init v
        selectors.Add obj_selectors
        Set obj_selectors = Nothing
    Next v
    Let text = CStr(obj.Item("text"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("selectors") = selectors
    dict("text") = text

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set selectors = New Collection
End Sub

