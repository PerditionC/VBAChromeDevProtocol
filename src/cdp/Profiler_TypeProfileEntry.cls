VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerTypeProfileEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Profiler.TypeProfileEntry
' Source offset and types for a parameter or return value.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' offset : integer
' Source offset of the parameter or end of function for return values.
Public offset AS integer

' types : array of cdpProfilerTypeObject
' The types for this parameter or return value.
Public types AS Collection

Public Function init(ByVal obj as Dictionary) As cdpProfilerTypeProfileEntry
    Dim v as Variant

    Let offset = CInt(obj.Item("offset"))
    For each v in obj.Item("types")
        Dim obj_types As cdpProfilerTypeObject: Set obj_types = New cdpProfilerTypeObject
        obj_types.init v
        types.Add obj_types
        Set obj_types = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("offset") = offset
    Set dict("types") = types

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set types = New Collection
End Sub
