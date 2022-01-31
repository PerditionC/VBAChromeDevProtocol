VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerTypeObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "TypeObject"
' Profiler.TypeObject
' Describes a type collected during runtime.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Name of a type collected with type profiling.
Public name AS string

Public Function init(ByVal obj as Dictionary) As cdpProfilerTypeObject
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name

    Set asDictionary = dict
End Function

