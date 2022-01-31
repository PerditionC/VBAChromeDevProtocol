VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpAccessibilityAXProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AXProperty"
' Accessibility.AXProperty
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : AXPropertyName
' The name of this property.
Public name AS string

' value : AXValue
' The value of this property.
Public value AS cdpAccessibilityAXValue

Public Function init(ByVal obj as Dictionary) As cdpAccessibilityAXProperty
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    Set value = New cdpAccessibilityAXValue
    value.init obj.Item("value")

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    Set dict("value") = value.asDictionary()

    Set asDictionary = dict
End Function

