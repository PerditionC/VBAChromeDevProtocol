VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpMemory"
' Memory [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Memory pressure level.
' PressureLevel AS string
Public Enum PressureLevel
    PL_moderate
    PL_critical
End Enum

' Heap profile sample.
' SamplingProfileNode AS object

' Array of heap profile samples.
' SamplingProfile AS object

' Executable module information
' Module AS object


Public Function PressureLevelToString(ByVal pl As PressureLevel) As String
    Dim retVal As String
    Select Case pl
        Case PL_moderate
            retVal = "moderate"
        Case PL_critical
            retVal = "critical"
        Case Else
            Debug.Print "Warning, unknown value " & pl
    End Select
    PressureLevelToString = retVal
End Function

Public Function StringToPressureLevel(ByVal s As String) As PressureLevel
    Dim retVal As PressureLevel
    Select Case s
        Case "moderate"
            retVal = PL_moderate
        Case "critical"
            retVal = PL_critical
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToPressureLevel = retVal
End Function


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
' Returns:  - 
Public Function getDOMCounters() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Memory.getDOMCounters", params)
    If cdp.ErrorCode = 0 Then
        Set getDOMCounters = results
    End If
End Function

' 
Public Sub prepareForLeakDetection()
    Dim params As New Dictionary
    cdp.InvokeMethod "Memory.prepareForLeakDetection", params
End Sub

' Simulate OomIntervention by purging V8 memory.
Public Sub forciblyPurgeJavaScriptMemory()
    Dim params As New Dictionary
    cdp.InvokeMethod "Memory.forciblyPurgeJavaScriptMemory", params
End Sub

' Enable/disable suppressing memory pressure notifications in all processes.
Public Sub setPressureNotificationsSuppressed( _
    ByVal suppressed AS boolean _
)
    ' suppressed: boolean If true, memory pressure notifications will be suppressed.

    Dim params As New Dictionary
    params("suppressed") = CBool(suppressed)

    cdp.InvokeMethod "Memory.setPressureNotificationsSuppressed", params
End Sub

' Simulate a memory pressure notification in all processes.
Public Sub simulatePressureNotification( _
    ByVal level AS string _
)
    ' level: PressureLevel Memory pressure level of the notification.

    Dim params As New Dictionary
    params("level") = CStr(level)

    cdp.InvokeMethod "Memory.simulatePressureNotification", params
End Sub

' Start collecting native memory profile.
Public Sub startSampling( _
    Optional ByVal samplingInterval AS Variant, _
    Optional ByVal suppressRandomness AS Variant _
)
    ' samplingInterval: integer(optional) Average number of bytes between samples.
    ' suppressRandomness: boolean(optional) Do not randomize intervals between samples.

    Dim params As New Dictionary
    If Not IsMissing(samplingInterval) Then params("samplingInterval") = CLng(samplingInterval)
    If Not IsMissing(suppressRandomness) Then params("suppressRandomness") = CBool(suppressRandomness)

    cdp.InvokeMethod "Memory.startSampling", params
End Sub

' Stop collecting native memory profile.
Public Sub stopSampling()
    Dim params As New Dictionary
    cdp.InvokeMethod "Memory.stopSampling", params
End Sub

' Retrieve native memory allocations profile
' collected since renderer process startup.
' Returns: profile - 
Public Function getAllTimeSamplingProfile() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Memory.getAllTimeSamplingProfile", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("profile") Then Set getAllTimeSamplingProfile = results("profile")
    End If
End Function

' Retrieve native memory allocations profile
' collected since browser process startup.
' Returns: profile - 
Public Function getBrowserSamplingProfile() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Memory.getBrowserSamplingProfile", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("profile") Then Set getBrowserSamplingProfile = results("profile")
    End If
End Function

' Retrieve native memory allocations profile collected since last
' `startSampling` call.
' Returns: profile - 
Public Function getSamplingProfile() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Memory.getSamplingProfile", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("profile") Then Set getSamplingProfile = results("profile")
    End If
End Function

