VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpHeapProfiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpHeapProfiler"
' HeapProfiler [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Heap snapshot object id.
' HeapSnapshotObjectId AS string

' Sampling Heap Profile node. Holds callsite information, allocation statistics and child nodes.
' SamplingHeapProfileNode AS object

' A single sample from a sampling profile.
' SamplingHeapProfileSample AS object

' Sampling profile.
' SamplingHeapProfile AS object



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

' Enables console to refer to the node with given id via $x (see Command Line API for more details
' $x functions).
Public Sub addInspectedHeapObject( _
    ByVal heapObjectId AS string _
)
    ' heapObjectId: HeapSnapshotObjectId Heap snapshot object id to be accessible by means of $x command line API.

    Dim params As New Dictionary
    params("heapObjectId") = CStr(heapObjectId)

    cdp.InvokeMethod "HeapProfiler.addInspectedHeapObject", params
End Sub

' 
Public Sub collectGarbage()
    Dim params As New Dictionary
    cdp.InvokeMethod "HeapProfiler.collectGarbage", params
End Sub

' 
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "HeapProfiler.disable", params
End Sub

' 
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "HeapProfiler.enable", params
End Sub

' 
' Returns: heapSnapshotObjectId - Id of the heap snapshot object corresponding to the passed remote object id.
Public Function getHeapObjectId( _
    ByVal objectId AS string _
) AS string
    ' objectId: Runtime.RemoteObjectId Identifier of the object to get heap object id for.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("HeapProfiler.getHeapObjectId", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("heapSnapshotObjectId") Then getHeapObjectId = results("heapSnapshotObjectId")
    End If
End Function

' 
' Returns: result - Evaluation result.
Public Function getObjectByHeapObjectId( _
    ByVal objectId AS string, _
    Optional ByVal objectGroup AS Variant _
) AS Dictionary
    ' objectId: HeapSnapshotObjectId 
    ' objectGroup: string(optional) Symbolic group name that can be used to release multiple objects.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)
    If Not IsMissing(objectGroup) Then params("objectGroup") = CStr(objectGroup)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("HeapProfiler.getObjectByHeapObjectId", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then Set getObjectByHeapObjectId = results("result")
    End If
End Function

' 
' Returns: profile - Return the sampling profile being collected.
Public Function getSamplingProfile() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("HeapProfiler.getSamplingProfile", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("profile") Then Set getSamplingProfile = results("profile")
    End If
End Function

' 
Public Sub startSampling( _
    Optional ByVal samplingInterval AS Variant _
)
    ' samplingInterval: number(optional) Average sample interval in bytes. Poisson distribution is used for the intervals. The
    '   default value is 32768 bytes.

    Dim params As New Dictionary
    If Not IsMissing(samplingInterval) Then params("samplingInterval") = CDbl(samplingInterval)

    cdp.InvokeMethod "HeapProfiler.startSampling", params
End Sub

' 
Public Sub startTrackingHeapObjects( _
    Optional ByVal trackAllocations AS Variant _
)
    ' trackAllocations: boolean(optional) 

    Dim params As New Dictionary
    If Not IsMissing(trackAllocations) Then params("trackAllocations") = CBool(trackAllocations)

    cdp.InvokeMethod "HeapProfiler.startTrackingHeapObjects", params
End Sub

' 
' Returns: profile - Recorded sampling heap profile.
Public Function stopSampling() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("HeapProfiler.stopSampling", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("profile") Then Set stopSampling = results("profile")
    End If
End Function

' 
Public Sub stopTrackingHeapObjects( _
    Optional ByVal reportProgress AS Variant, _
    Optional ByVal treatGlobalObjectsAsRoots AS Variant, _
    Optional ByVal captureNumericValue AS Variant _
)
    ' reportProgress: boolean(optional) If true 'reportHeapSnapshotProgress' events will be generated while snapshot is being taken
    '   when the tracking is stopped.
    ' treatGlobalObjectsAsRoots: boolean(optional) 
    ' captureNumericValue: boolean(optional) If true, numerical values are included in the snapshot

    Dim params As New Dictionary
    If Not IsMissing(reportProgress) Then params("reportProgress") = CBool(reportProgress)
    If Not IsMissing(treatGlobalObjectsAsRoots) Then params("treatGlobalObjectsAsRoots") = CBool(treatGlobalObjectsAsRoots)
    If Not IsMissing(captureNumericValue) Then params("captureNumericValue") = CBool(captureNumericValue)

    cdp.InvokeMethod "HeapProfiler.stopTrackingHeapObjects", params
End Sub

' 
Public Sub takeHeapSnapshot( _
    Optional ByVal reportProgress AS Variant, _
    Optional ByVal treatGlobalObjectsAsRoots AS Variant, _
    Optional ByVal captureNumericValue AS Variant _
)
    ' reportProgress: boolean(optional) If true 'reportHeapSnapshotProgress' events will be generated while snapshot is being taken.
    ' treatGlobalObjectsAsRoots: boolean(optional) If true, a raw snapshot without artificial roots will be generated
    ' captureNumericValue: boolean(optional) If true, numerical values are included in the snapshot

    Dim params As New Dictionary
    If Not IsMissing(reportProgress) Then params("reportProgress") = CBool(reportProgress)
    If Not IsMissing(treatGlobalObjectsAsRoots) Then params("treatGlobalObjectsAsRoots") = CBool(treatGlobalObjectsAsRoots)
    If Not IsMissing(captureNumericValue) Then params("captureNumericValue") = CBool(captureNumericValue)

    cdp.InvokeMethod "HeapProfiler.takeHeapSnapshot", params
End Sub

