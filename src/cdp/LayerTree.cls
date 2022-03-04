VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpLayerTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpLayerTree"
' LayerTree [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Unique Layer identifier.
' LayerId AS string

' Unique snapshot identifier.
' SnapshotId AS string

' Rectangle where scrolling happens on the main thread.
' ScrollRect AS object

' Sticky position constraints.
' StickyPositionConstraint AS object

' Serialized fragment of layer picture along with its offset within the layer.
' PictureTile AS object

' Information about a compositing layer.
' Layer AS object

' Array of timings, one per paint step.
' PaintProfile AS array



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

' Provides the reasons why the given layer was composited.
' Returns:  - 
Public Function compositingReasons( _
    ByVal layerId AS string _
) AS Dictionary
    ' layerId: LayerId The id of the layer for which we want to get the reasons it was composited.

    Dim params As New Dictionary
    params("layerId") = CStr(layerId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("LayerTree.compositingReasons", params)
    If cdp.ErrorCode = 0 Then
        Set compositingReasons = results
    End If
End Function

' Disables compositing tree inspection.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "LayerTree.disable", params
End Sub

' Enables compositing tree inspection.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "LayerTree.enable", params
End Sub

' Returns the snapshot identifier.
' Returns: snapshotId - The id of the snapshot.
Public Function loadSnapshot( _
    ByVal tiles AS Collection _
) AS string
    ' tiles: array An array of tiles composing the snapshot.

    Dim params As New Dictionary
    Set params("tiles") = tiles

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("LayerTree.loadSnapshot", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("snapshotId") Then loadSnapshot = results("snapshotId")
    End If
End Function

' Returns the layer snapshot identifier.
' Returns: snapshotId - The id of the layer snapshot.
Public Function makeSnapshot( _
    ByVal layerId AS string _
) AS string
    ' layerId: LayerId The id of the layer.

    Dim params As New Dictionary
    params("layerId") = CStr(layerId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("LayerTree.makeSnapshot", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("snapshotId") Then makeSnapshot = results("snapshotId")
    End If
End Function

' 
' Returns: timings - The array of paint profiles, one per run.
Public Function profileSnapshot( _
    ByVal snapshotId AS string, _
    Optional ByVal minRepeatCount AS Variant, _
    Optional ByVal minDuration AS Variant, _
    Optional ByVal clipRect AS Variant _
) AS Collection
    ' snapshotId: SnapshotId The id of the layer snapshot.
    ' minRepeatCount: integer(optional) The maximum number of times to replay the snapshot (1, if not specified).
    ' minDuration: number(optional) The minimum duration (in seconds) to replay the snapshot.
    ' clipRect: DOM.Rect(optional) The clip rectangle to apply when replaying the snapshot.

    Dim params As New Dictionary
    params("snapshotId") = CStr(snapshotId)
    If Not IsMissing(minRepeatCount) Then params("minRepeatCount") = CLng(minRepeatCount)
    If Not IsMissing(minDuration) Then params("minDuration") = CDbl(minDuration)
    If Not IsMissing(clipRect) Then Set params("clipRect") = clipRect

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("LayerTree.profileSnapshot", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("timings") Then Set profileSnapshot = results("timings")
    End If
End Function

' Releases layer snapshot captured by the back-end.
Public Sub releaseSnapshot( _
    ByVal snapshotId AS string _
)
    ' snapshotId: SnapshotId The id of the layer snapshot.

    Dim params As New Dictionary
    params("snapshotId") = CStr(snapshotId)

    cdp.InvokeMethod "LayerTree.releaseSnapshot", params
End Sub

' Replays the layer snapshot and returns the resulting bitmap.
' Returns: dataURL - A data: URL for resulting image.
Public Function replaySnapshot( _
    ByVal snapshotId AS string, _
    Optional ByVal fromStep AS Variant, _
    Optional ByVal toStep AS Variant, _
    Optional ByVal pScale AS Variant _
) AS string
    ' snapshotId: SnapshotId The id of the layer snapshot.
    ' fromStep: integer(optional) The first step to replay from (replay from the very start if not specified).
    ' toStep: integer(optional) The last step to replay to (replay till the end if not specified).
    ' pScale: number(optional) The scale to apply while replaying (defaults to 1).

    Dim params As New Dictionary
    params("snapshotId") = CStr(snapshotId)
    If Not IsMissing(fromStep) Then params("fromStep") = CLng(fromStep)
    If Not IsMissing(toStep) Then params("toStep") = CLng(toStep)
    If Not IsMissing(pScale) Then params("scale") = CDbl(pScale)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("LayerTree.replaySnapshot", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("dataURL") Then replaySnapshot = results("dataURL")
    End If
End Function

' Replays the layer snapshot and returns canvas log.
' Returns: commandLog - The array of canvas function calls.
Public Function snapshotCommandLog( _
    ByVal snapshotId AS string _
) AS Collection
    ' snapshotId: SnapshotId The id of the layer snapshot.

    Dim params As New Dictionary
    params("snapshotId") = CStr(snapshotId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("LayerTree.snapshotCommandLog", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("commandLog") Then Set snapshotCommandLog = results("commandLog")
    End If
End Function

