VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerScriptCoverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScriptCoverage"
' Profiler.ScriptCoverage
' Coverage data for a JavaScript script.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' scriptId : Runtime.ScriptId
' JavaScript script id.
Public scriptId AS string

' url : string
' JavaScript script name or url.
Public url AS string

' functions : array of cdpProfilerFunctionCoverage
' Functions contained in the script that has coverage data.
Public functions AS Collection

Public Function init(ByVal obj as Dictionary) As cdpProfilerScriptCoverage
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let scriptId = CStr(obj.Item("scriptId"))
    Let url = CStr(obj.Item("url"))
    For each v in obj.Item("functions")
        Dim obj_functions As cdpProfilerFunctionCoverage: Set obj_functions = New cdpProfilerFunctionCoverage
        obj_functions.init v
        functions.Add obj_functions
        Set obj_functions = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("scriptId") = scriptId
    dict("url") = url
    Set dict("functions") = functions

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set functions = New Collection
End Sub

