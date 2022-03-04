VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpProfiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpProfiler"
' Profiler
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Profile node. Holds callsite information, execution statistics and child nodes.
' ProfileNode AS object

' Profile.
' Profile AS object

' Specifies a number of samples attributed to a certain source position.
' PositionTickInfo AS object

' Coverage data for a source range.
' CoverageRange AS object

' Coverage data for a JavaScript function.
' FunctionCoverage AS object

' Coverage data for a JavaScript script.
' ScriptCoverage AS object

' Describes a type collected during runtime.
' TypeObject AS object

' Source offset and types for a parameter or return value.
' TypeProfileEntry AS object

' Type profile data collected during runtime for a JavaScript script.
' ScriptTypeProfile AS object



Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' 
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Profiler.disable", params
End Sub

' 
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Profiler.enable", params
End Sub

' Collect coverage data for the current isolate. The coverage data may be incomplete due to
' garbage collection.
' Returns: result - Coverage data for the current isolate.
Public Function getBestEffortCoverage() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Profiler.getBestEffortCoverage", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then Set getBestEffortCoverage = results("result")
    End If
End Function

' Changes CPU profiler sampling interval. Must be called before CPU profiles recording started.
Public Sub setSamplingInterval( _
    ByVal interval AS Long _
)
    ' interval: integer New sampling interval in microseconds.

    Dim params As New Dictionary
    params("interval") = CLng(interval)

    cdp.InvokeMethod "Profiler.setSamplingInterval", params
End Sub

' 
Public Sub start()
    Dim params As New Dictionary
    cdp.InvokeMethod "Profiler.start", params
End Sub

' Enable precise code coverage. Coverage data for JavaScript executed before enabling precise code
' coverage may be incomplete. Enabling prevents running optimized code and resets execution
' counters.
' Returns: timestamp - Monotonically increasing time (in seconds) when the coverage update was taken in the backend.
Public Function startPreciseCoverage( _
    Optional ByVal callCount AS Variant, _
    Optional ByVal detailed AS Variant, _
    Optional ByVal allowTriggeredUpdates AS Variant _
) AS Double
    ' callCount: boolean(optional) Collect accurate call counts beyond simple 'covered' or 'not covered'.
    ' detailed: boolean(optional) Collect block-based coverage.
    ' allowTriggeredUpdates: boolean(optional) Allow the backend to send updates on its own initiative

    Dim params As New Dictionary
    If Not IsMissing(callCount) Then params("callCount") = CBool(callCount)
    If Not IsMissing(detailed) Then params("detailed") = CBool(detailed)
    If Not IsMissing(allowTriggeredUpdates) Then params("allowTriggeredUpdates") = CBool(allowTriggeredUpdates)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Profiler.startPreciseCoverage", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("timestamp") Then startPreciseCoverage = results("timestamp")
    End If
End Function

' Enable type profile.
' Experimental
Public Sub startTypeProfile()
    Dim params As New Dictionary
    cdp.InvokeMethod "Profiler.startTypeProfile", params
End Sub

' 
' Returns: profile - Recorded profile.
Public Function stopProfiler() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Profiler.stop", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("profile") Then Set stopProfiler = results("profile")
    End If
End Function

' Disable precise code coverage. Disabling releases unnecessary execution count records and allows
' executing optimized code.
Public Sub stopPreciseCoverage()
    Dim params As New Dictionary
    cdp.InvokeMethod "Profiler.stopPreciseCoverage", params
End Sub

' Disable type profile. Disabling releases type profile data collected so far.
' Experimental
Public Sub stopTypeProfile()
    Dim params As New Dictionary
    cdp.InvokeMethod "Profiler.stopTypeProfile", params
End Sub

' Collect coverage data for the current isolate, and resets execution counters. Precise code
' coverage needs to have started.
' Returns:  - 
Public Function takePreciseCoverage() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Profiler.takePreciseCoverage", params)
    If cdp.ErrorCode = 0 Then
        Set takePreciseCoverage = results
    End If
End Function

' Collect type profile.
' Returns: result - Type profile for all scripts since startTypeProfile() was turned on.
' Experimental
Public Function takeTypeProfile() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Profiler.takeTypeProfile", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then Set takeTypeProfile = results("result")
    End If
End Function

