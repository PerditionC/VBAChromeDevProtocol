VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpTracing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpTracing"
' Tracing [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Configuration for memory dump. Used only when "memory-infra" category is enabled.
' MemoryDumpConfig AS object

' 
' TraceConfig AS object

' Data format of a trace. Can be either the legacy JSON format or the
'   protocol buffer format. Note that the JSON format will be deprecated soon.
' StreamFormat AS string
Public Enum StreamFormat
    SF_json
    SF_proto
End Enum

' Compression type to use for traces returned via streams.
' StreamCompression AS string
Public Enum StreamCompression
    SC_none
    SC_gzip
End Enum

' Details exposed when memory request explicitly declared.
'   Keep consistent with memory_dump_request_args.h and
'   memory_instrumentation.mojom
' MemoryDumpLevelOfDetail AS string
Public Enum MemoryDumpLevelOfDetail
    MDLOD_background
    MDLOD_light
    MDLOD_detailed
End Enum

' Backend type to use for tracing. `chrome` uses the Chrome-integrated
'   tracing service and is supported on all platforms. `system` is only
'   supported on Chrome OS and uses the Perfetto system tracing service.
'   `auto` chooses `system` when the perfettoConfig provided to Tracing.start
'   specifies at least one non-Chrome data source; otherwise uses `chrome`.
' TracingBackend AS string
Public Enum TracingBackend
    TB_auto
    TB_chrome
    TB_system
End Enum


Public Function StreamFormatToString(ByVal sf As StreamFormat) As String
    Dim retVal As String
    Select Case sf
        Case SF_json
            retVal = "json"
        Case SF_proto
            retVal = "proto"
        Case Else
            Debug.Print "Warning, unknown value " & sf
    End Select
    StreamFormatToString = retVal
End Function

Public Function StringToStreamFormat(ByVal s As String) As StreamFormat
    Dim retVal As StreamFormat
    Select Case s
        Case "json"
            retVal = SF_json
        Case "proto"
            retVal = SF_proto
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToStreamFormat = retVal
End Function


Public Function StreamCompressionToString(ByVal sc As StreamCompression) As String
    Dim retVal As String
    Select Case sc
        Case SC_none
            retVal = "none"
        Case SC_gzip
            retVal = "gzip"
        Case Else
            Debug.Print "Warning, unknown value " & sc
    End Select
    StreamCompressionToString = retVal
End Function

Public Function StringToStreamCompression(ByVal s As String) As StreamCompression
    Dim retVal As StreamCompression
    Select Case s
        Case "none"
            retVal = SC_none
        Case "gzip"
            retVal = SC_gzip
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToStreamCompression = retVal
End Function


Public Function MemoryDumpLevelOfDetailToString(ByVal mdlod As MemoryDumpLevelOfDetail) As String
    Dim retVal As String
    Select Case mdlod
        Case MDLOD_background
            retVal = "background"
        Case MDLOD_light
            retVal = "light"
        Case MDLOD_detailed
            retVal = "detailed"
        Case Else
            Debug.Print "Warning, unknown value " & mdlod
    End Select
    MemoryDumpLevelOfDetailToString = retVal
End Function

Public Function StringToMemoryDumpLevelOfDetail(ByVal s As String) As MemoryDumpLevelOfDetail
    Dim retVal As MemoryDumpLevelOfDetail
    Select Case s
        Case "background"
            retVal = MDLOD_background
        Case "light"
            retVal = MDLOD_light
        Case "detailed"
            retVal = MDLOD_detailed
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToMemoryDumpLevelOfDetail = retVal
End Function


Public Function TracingBackendToString(ByVal tb As TracingBackend) As String
    Dim retVal As String
    Select Case tb
        Case TB_auto
            retVal = "auto"
        Case TB_chrome
            retVal = "chrome"
        Case TB_system
            retVal = "system"
        Case Else
            Debug.Print "Warning, unknown value " & tb
    End Select
    TracingBackendToString = retVal
End Function

Public Function StringToTracingBackend(ByVal s As String) As TracingBackend
    Dim retVal As TracingBackend
    Select Case s
        Case "auto"
            retVal = TB_auto
        Case "chrome"
            retVal = TB_chrome
        Case "system"
            retVal = TB_system
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToTracingBackend = retVal
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

' Stop trace events collection.
Public Sub pEnd()
    Dim params As New Dictionary
    cdp.InvokeMethod "Tracing.end", params
End Sub

' Gets supported tracing categories.
' Returns: categories - A list of supported tracing categories.
Public Function getCategories() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Tracing.getCategories", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("categories") Then Set getCategories = results("categories")
    End If
End Function

' Record a clock sync marker in the trace.
Public Sub recordClockSyncMarker( _
    ByVal syncId AS string _
)
    ' syncId: string The ID of this clock sync marker

    Dim params As New Dictionary
    params("syncId") = CStr(syncId)

    cdp.InvokeMethod "Tracing.recordClockSyncMarker", params
End Sub

' Request a global memory dump.
' Returns:  - 
Public Function requestMemoryDump( _
    Optional ByVal deterministic AS Variant, _
    Optional ByVal levelOfDetail AS Variant _
) AS Dictionary
    ' deterministic: boolean(optional) Enables more deterministic results by forcing garbage collection
    ' levelOfDetail: MemoryDumpLevelOfDetail(optional) Specifies level of details in memory dump. Defaults to "detailed".

    Dim params As New Dictionary
    If Not IsMissing(deterministic) Then params("deterministic") = CBool(deterministic)
    If Not IsMissing(levelOfDetail) Then params("levelOfDetail") = CStr(levelOfDetail)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Tracing.requestMemoryDump", params)
    If cdp.ErrorCode = 0 Then
        Set requestMemoryDump = results
    End If
End Function

' Start trace events collection.
Public Sub start( _
    Optional ByVal categories AS Variant, _
    Optional ByVal options AS Variant, _
    Optional ByVal bufferUsageReportingInterval AS Variant, _
    Optional ByVal transferMode AS Variant, _
    Optional ByVal streamFormat AS Variant, _
    Optional ByVal streamCompression AS Variant, _
    Optional ByVal traceConfig AS Variant, _
    Optional ByVal perfettoConfig AS Variant, _
    Optional ByVal tracingBackend AS Variant _
)
    ' categories: string(optional) Category/tag filter
    ' options: string(optional) Tracing options
    ' bufferUsageReportingInterval: number(optional) If set, the agent will issue bufferUsage events at this interval, specified in milliseconds
    ' transferMode: string(optional) Whether to report trace events as series of dataCollected events or to save trace to a
    '   stream (defaults to `ReportEvents`).
    ' streamFormat: StreamFormat(optional) Trace data format to use. This only applies when using `ReturnAsStream`
    '   transfer mode (defaults to `json`).
    ' streamCompression: StreamCompression(optional) Compression format to use. This only applies when using `ReturnAsStream`
    '   transfer mode (defaults to `none`)
    ' traceConfig: TraceConfig(optional) 
    ' perfettoConfig: binary(optional) Base64-encoded serialized perfetto.protos.TraceConfig protobuf message
    '   When specified, the parameters `categories`, `options`, `traceConfig`
    '   are ignored.
    ' tracingBackend: TracingBackend(optional) Backend type (defaults to `auto`)

    Dim params As New Dictionary
    If Not IsMissing(categories) Then params("categories") = CStr(categories)
    If Not IsMissing(options) Then params("options") = CStr(options)
    If Not IsMissing(bufferUsageReportingInterval) Then params("bufferUsageReportingInterval") = CDbl(bufferUsageReportingInterval)
    If Not IsMissing(transferMode) Then params("transferMode") = CStr(transferMode)
    If Not IsMissing(streamFormat) Then params("streamFormat") = CStr(streamFormat)
    If Not IsMissing(streamCompression) Then params("streamCompression") = CStr(streamCompression)
    If Not IsMissing(traceConfig) Then Set params("traceConfig") = traceConfig
    If Not IsMissing(perfettoConfig) Then params("perfettoConfig") = CStr(perfettoConfig)
    If Not IsMissing(tracingBackend) Then params("tracingBackend") = CStr(tracingBackend)

    cdp.InvokeMethod "Tracing.start", params
End Sub

