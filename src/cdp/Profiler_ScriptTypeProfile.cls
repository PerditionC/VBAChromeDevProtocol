VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerScriptTypeProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScriptTypeProfile"
' Profiler.ScriptTypeProfile
' Type profile data collected during runtime for a JavaScript script.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' scriptId : Runtime.ScriptId
' JavaScript script id.
Public scriptId AS string

' url : string
' JavaScript script name or url.
Public url AS string

' entries : array of cdpProfilerTypeProfileEntry
' Type profile entries for parameters and return values of the functions in the script.
Public entries AS Collection

Public Function init(ByVal obj as Dictionary) As cdpProfilerScriptTypeProfile
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let scriptId = CStr(obj.Item("scriptId"))
    Let url = CStr(obj.Item("url"))
    For each v in obj.Item("entries")
        Dim obj_entries As cdpProfilerTypeProfileEntry: Set obj_entries = New cdpProfilerTypeProfileEntry
        obj_entries.init v
        entries.Add obj_entries
        Set obj_entries = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("scriptId") = scriptId
    dict("url") = url
    Set dict("entries") = entries

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set entries = New Collection
End Sub

