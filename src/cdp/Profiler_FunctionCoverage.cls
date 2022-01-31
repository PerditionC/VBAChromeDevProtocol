VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfilerFunctionCoverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FunctionCoverage"
' Profiler.FunctionCoverage
' Coverage data for a JavaScript function.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' functionName : string
' JavaScript function name.
Public functionName AS string

' ranges : array of cdpProfilerCoverageRange
' Source ranges inside the function with coverage data.
Public ranges AS Collection

' isBlockCoverage : boolean
' Whether coverage data for this function has block granularity.
Public isBlockCoverage AS boolean

Public Function init(ByVal obj as Dictionary) As cdpProfilerFunctionCoverage
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let functionName = CStr(obj.Item("functionName"))
    For each v in obj.Item("ranges")
        Dim obj_ranges As cdpProfilerCoverageRange: Set obj_ranges = New cdpProfilerCoverageRange
        obj_ranges.init v
        ranges.Add obj_ranges
        Set obj_ranges = Nothing
    Next v
    Let isBlockCoverage = CBool(obj.Item("isBlockCoverage"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("functionName") = functionName
    Set dict("ranges") = ranges
    dict("isBlockCoverage") = isBlockCoverage

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set ranges = New Collection
End Sub

